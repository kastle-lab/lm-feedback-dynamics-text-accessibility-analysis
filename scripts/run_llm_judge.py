#!/usr/bin/env python3
"""Simple Ollama LLM-as-judge runner for result CSV rows."""

import argparse
import csv
import importlib
import json
import os
import sys
from pathlib import Path

from ollama_interface import chat_with_model

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

DEFAULT_INPUTS = [
    "results",
    "results/local_incremental",
    "results/local_holistic",
]


class BlankDict(dict):
    def __missing__(self, key):
        return ""


def read_file(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def load_prompt_module(module_name):
    prompt_module = importlib.import_module(module_name)
    prompt_module = importlib.reload(prompt_module)

    if not hasattr(prompt_module, "SYSTEM_PROMPT"):
        raise ValueError(f"{module_name} does not define SYSTEM_PROMPT")
    if not hasattr(prompt_module, "USER_PROMPT_TEMPLATE"):
        raise ValueError(f"{module_name} does not define USER_PROMPT_TEMPLATE")

    return prompt_module.SYSTEM_PROMPT, prompt_module.USER_PROMPT_TEMPLATE


def load_original_chapters(path):
    with open(path, "r", encoding="utf-8") as f:
        chapters = json.load(f)

    return {
        str(chapter.get("header", "")): str(chapter.get("text", ""))
        for chapter in chapters
        if chapter.get("header")
    }


def append_jsonl(path, record):
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def append_tracker(path, row):
    fieldnames = [
        "source_csv",
        "row_number",
        "judge_model",
        "generation_model",
        "chapter_key",
        "iteration",
        "status",
    ]
    new_file = not path.exists()

    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if new_file:
            writer.writeheader()
        writer.writerow({key: row.get(key, "") for key in fieldnames})


def infer_generation_model(csv_path):
    name = csv_path.name
    if name.endswith("_agent_log.csv"):
        return name[: -len("_agent_log.csv")]
    return ""


def resolve_path(path_value):
    path = Path(path_value)
    if path.exists():
        return path

    project_path = PROJECT_ROOT / path_value
    if project_path.exists():
        return project_path

    return path


def discover_input_csvs(inputs):
    csv_paths = []
    for item in inputs:
        path = resolve_path(item)
        if path.is_dir():
            csv_paths.extend(sorted(path.glob("*_agent_log*.csv")))
        elif path.is_file():
            csv_paths.append(path)
        else:
            raise FileNotFoundError(f"Input path does not exist: {path}")

    return csv_paths


def format_prompt(template, values):
    values = BlankDict({key: "" if value is None else str(value) for key, value in values.items()})
    return template.format_map(values)


def judge_row(judge_model, system_prompt, user_prompt, temperature):
    response = chat_with_model(
        model_name=judge_model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        options={"temperature": temperature},
    )
    return response["message"]["content"]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "inputs",
        nargs="*",
        default=None,
        help="Optional CSV files or directories. Defaults to all local result logs.",
    )
    parser.add_argument("--judge-model", default=os.environ.get("OLLAMA_JUDGE_MODEL"))
    parser.add_argument("--prompt-module", default="prompts.llm_judge_0_shot")
    parser.add_argument("--original-json", default="data/chapters.json")
    parser.add_argument("--output-jsonl", default="results/llm_judge/judge_results.jsonl")
    parser.add_argument("--tracker-csv", default="results/llm_judge/judge_tracker.csv")
    parser.add_argument("--temperature", type=float, default=0)
    parser.add_argument("--limit", type=int, default=None)
    args = parser.parse_args()

    if not args.judge_model:
        raise SystemExit("Set --judge-model or OLLAMA_JUDGE_MODEL.")

    output_jsonl = Path(args.output_jsonl)
    tracker_csv = Path(args.tracker_csv)
    output_jsonl.parent.mkdir(parents=True, exist_ok=True)
    tracker_csv.parent.mkdir(parents=True, exist_ok=True)

    system_prompt, user_template = load_prompt_module(args.prompt_module)
    original_chapters = load_original_chapters(resolve_path(args.original_json))
    input_csvs = discover_input_csvs(args.inputs or DEFAULT_INPUTS)
    if not input_csvs:
        raise SystemExit("No *_agent_log.csv files found.")

    print(f"Found {len(input_csvs)} input CSV files.")
    judged_count = 0

    for input_csv in input_csvs:
        generation_model = infer_generation_model(input_csv)

        with open(input_csv, "r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)

            for row_number, row in enumerate(reader, start=1):
                if args.limit is not None and judged_count >= args.limit:
                    break

                values = dict(row)
                values["source_csv"] = str(input_csv)
                values["row_number"] = row_number
                values["generation_model"] = row.get("model") or generation_model
                values["original_chapter"] = original_chapters.get(row.get("chapter_key", ""), "")
                values["Insert_original_chapter_here"] = values["original_chapter"]
                values["Insert_rewritten_chapter_here"] = row.get("text", "")

                print(
                    f"Judging {judged_count + 1}: "
                    f"{input_csv} row {row_number} "
                    f"{values.get('generation_model', '')} "
                    f"{values.get('chapter_key', '')} "
                    f"iteration {values.get('iteration', '')}"
                )

                try:
                    user_prompt = format_prompt(user_template, values)
                    judgment = judge_row(
                        judge_model=args.judge_model,
                        system_prompt=system_prompt,
                        user_prompt=user_prompt,
                        temperature=args.temperature,
                    )
                    status = "ok"
                    error = ""
                except Exception as exc:
                    judgment = ""
                    status = "error"
                    error = str(exc)

                append_jsonl(
                    output_jsonl,
                    {
                        "source_csv": str(input_csv),
                        "row_number": row_number,
                        "judge_model": args.judge_model,
                        "generation_model": values.get("generation_model", ""),
                        "source_row": row,
                        "judgment": judgment,
                        "status": status,
                        "error": error,
                    },
                )
                append_tracker(
                    tracker_csv,
                    {
                        "source_csv": str(input_csv),
                        "row_number": row_number,
                        "judge_model": args.judge_model,
                        "generation_model": values.get("generation_model", ""),
                        "chapter_key": values.get("chapter_key", ""),
                        "iteration": values.get("iteration", ""),
                        "status": status,
                    },
                )
                judged_count += 1

        if args.limit is not None and judged_count >= args.limit:
            break

    print(f"Judged {judged_count} rows.")
    print(f"Wrote judgments to {output_jsonl}")
    print(f"Wrote tracker to {tracker_csv}")


if __name__ == "__main__":
    if Path.cwd().name != "AIED 2026":
        print("Tip: run this from the AIED 2026 directory.", file=sys.stderr)
    main()
