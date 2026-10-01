# Summarize per-model CSV logs from the Ollama experiment

import argparse
import ast
import csv
import re
from collections import Counter
from pathlib import Path


def as_bool(value):
    return str(value).strip().lower() == "true"


def as_float(value):
    if value in ("", None):
        return None
    try:
        return float(value)
    except ValueError:
        return None


def average(values):
    values = [value for value in values if value is not None]
    if not values:
        return ""
    return sum(values) / len(values)


def labels(value):
    if not value:
        return []
    try:
        parsed = ast.literal_eval(value)
        if isinstance(parsed, list):
            return [str(item) for item in parsed]
    except (SyntaxError, ValueError):
        pass
    return [item.strip() for item in str(value).split(",") if item.strip()]


def dystext_parts(value):
    numbers = re.findall(r"-?\d+(?:\.\d+)?", str(value or ""))
    if len(numbers) < 2:
        return None, None, as_float(value)
    visual = float(numbers[0])
    content = float(numbers[-1])
    return visual, content, visual + content


def summarize_log(log_path):
    log_path = Path(log_path)

    with open(log_path, "r", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    if not rows:
        return None

    failed_checks = Counter()
    for row in rows:
        failed_checks.update(labels(row.get("failed_checks")))

    reached_judge = [row for row in rows if as_bool(row.get("reached_judge"))]
    chapters = {row["chapter_key"] for row in rows if row.get("chapter_key")}
    dystext_values = [dystext_parts(row.get("dystext_score")) for row in reached_judge]

    return {
        "model": rows[0].get("model") or log_path.name.replace("_agent_log.csv", "").replace(".csv", ""),
        "chapters": len(chapters),
        "iterations": len(rows),
        "checker_passed_iterations": len(reached_judge),
        "checker_pass_rate": len(reached_judge) / len(rows) if rows else "",
        "longphrase_failures": failed_checks.get("longphrase", 0),
        "factual_drift_failures": failed_checks.get("factual_drift", 0),
        "avg_lix_score": average(as_float(row.get("lix_score")) for row in reached_judge),
        "avg_dystext_visual": average(value[0] for value in dystext_values),
        "avg_dystext_content": average(value[1] for value in dystext_values),
        "avg_dystext_score": average(value[2] for value in dystext_values),
        "checker_failed_labels": ";".join(
            f"{label}:{count}" for label, count in sorted(failed_checks.items())
        ),
    }


def write_summary(path, rows):
    if not rows:
        raise SystemExit("No experiment logs found.")

    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("results_dir", nargs="?", default="results/ollama")
    args = parser.parse_args()

    results_dir = Path(args.results_dir)
    summaries = [
        summary
        for summary in (
            summarize_log(path)
            for path in sorted(results_dir.glob("*_agent_log.csv"))
        )
        if summary
    ]
    write_summary(results_dir / "summary.csv", summaries)