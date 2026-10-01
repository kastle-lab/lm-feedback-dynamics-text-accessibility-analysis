# Consolidate LLM judge JSONL outputs by unique source row

import argparse
import csv
import json
from collections import Counter
from pathlib import Path


def load_allowed_rows(input_paths):
    allowed = set()
    for path in input_paths:
        if path.is_dir():
            files = sorted(path.glob("*_agent_log*.csv"))
        else:
            files = [path]

        for csv_path in files:
            if "llm_judge" in csv_path.resolve().parts:
                continue
            with open(csv_path, "r", newline="", encoding="utf-8") as f:
                for row_number, _ in enumerate(csv.DictReader(f), start=1):
                    allowed.add((str(csv_path), str(row_number)))
                    allowed.add((csv_path.name, str(row_number)))
    return allowed


def source_key(item):
    return (
        str(item.get("source_csv", "")),
        str(item.get("row_number", "")),
        str(item.get("judge_model", "")),
    )


def short_source_key(item):
    return (
        Path(str(item.get("source_csv", ""))).name,
        str(item.get("row_number", "")),
        str(item.get("judge_model", "")),
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("inputs", nargs="+", type=Path, help="Judge result JSONL files to consolidate.")
    parser.add_argument("--output-jsonl", type=Path, required=True)
    parser.add_argument("--tracker-csv", type=Path, required=True)
    parser.add_argument(
        "--allowed-input",
        nargs="*",
        type=Path,
        default=[
            Path("results"),
            Path("results/local_incremental"),
            Path("results/local_holistic"),
        ],
        help="Original result CSV files/directories used to define valid source rows.",
    )
    args = parser.parse_args()

    allowed = load_allowed_rows(args.allowed_input)
    seen = set()
    source_counts = Counter()
    kept = 0
    skipped_duplicate = 0
    skipped_invalid_source = 0
    skipped_json_error = 0

    args.output_jsonl.parent.mkdir(parents=True, exist_ok=True)
    args.tracker_csv.parent.mkdir(parents=True, exist_ok=True)

    with open(args.output_jsonl, "w", encoding="utf-8") as out_f, open(
        args.tracker_csv, "w", newline="", encoding="utf-8"
    ) as tracker_f:
        tracker = csv.DictWriter(
            tracker_f,
            fieldnames=[
                "source_csv",
                "row_number",
                "judge_model",
                "generation_model",
                "chapter_key",
                "iteration",
                "status",
            ],
        )
        tracker.writeheader()

        for input_path in args.inputs:
            with open(input_path, "r", encoding="utf-8") as f:
                for line in f:
                    if not line.strip():
                        continue
                    try:
                        item = json.loads(line)
                    except json.JSONDecodeError:
                        skipped_json_error += 1
                        continue

                    allowed_key = (str(item.get("source_csv", "")), str(item.get("row_number", "")))
                    allowed_short_key = (Path(str(item.get("source_csv", ""))).name, str(item.get("row_number", "")))
                    if allowed_key not in allowed and allowed_short_key not in allowed:
                        skipped_invalid_source += 1
                        continue

                    key = source_key(item)
                    short_key = short_source_key(item)
                    if key in seen or short_key in seen:
                        skipped_duplicate += 1
                        continue

                    seen.add(key)
                    seen.add(short_key)
                    source_counts[item.get("source_csv", "")] += 1
                    out_f.write(json.dumps(item, ensure_ascii=False) + "\n")
                    tracker.writerow(
                        {
                            "source_csv": item.get("source_csv", ""),
                            "row_number": item.get("row_number", ""),
                            "judge_model": item.get("judge_model", ""),
                            "generation_model": item.get("generation_model", ""),
                            "chapter_key": item.get("chapter_key", ""),
                            "iteration": item.get("iteration", ""),
                            "status": item.get("status", ""),
                        }
                    )
                    kept += 1

    print(f"Allowed source rows: {len(allowed)}")
    print(f"Kept: {kept}")
    print(f"Skipped duplicates: {skipped_duplicate}")
    print(f"Skipped invalid source rows: {skipped_invalid_source}")
    print(f"Skipped JSON errors: {skipped_json_error}")
    print(f"Wrote: {args.output_jsonl}")
    print(f"Wrote: {args.tracker_csv}")
    print("\nKept rows by source:")
    for source, count in source_counts.most_common():
        print(f"{count:6d} {source}")