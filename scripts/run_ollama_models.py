import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path


DEFAULT_CONFIG = "config/ollama_models.json"


def safe_name(model_name):
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", model_name).strip("_")


def load_config(path):
    with open(path, "r", encoding="utf-8") as f:
        config = json.load(f)

    required = ["input_json", "results_dir", "models"]
    missing = [key for key in required if key not in config]
    if missing:
        raise ValueError(f"Missing config keys: {', '.join(missing)}")

    if not isinstance(config["models"], list) or not config["models"]:
        raise ValueError("Config must include at least one model.")

    return config


def check_paths(config):
    paths = [
        Path(config["input_json"]),
        Path("scripts/ollama_interface.py"),
    ]
    missing = [str(path) for path in paths if not path.exists()]
    if missing:
        raise SystemExit(
            "Missing required files. Run this command from the AIED 2026 directory. "
            f"Missing: {', '.join(missing)}"
        )


def make_limited_input(input_json, results_dir, limit):
    if limit is None:
        return input_json

    with open(input_json, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, list):
        raise ValueError("Input JSON must be a list of chapter objects.")

    limited = data[:limit]
    results_dir.mkdir(parents=True, exist_ok=True)
    limited_path = results_dir / f"_limited_{limit}_chapters.json"
    with open(limited_path, "w", encoding="utf-8") as f:
        json.dump(limited, f, ensure_ascii=False, indent=2)

    return limited_path


def run_model(model, input_json, results_dir, temperature, max_passes):
    model_name = str(model)
    file_stem = safe_name(model_name)

    output_json = results_dir / f"{file_stem}_output.json"
    log_csv = results_dir / f"{file_stem}_agent_log.csv"

    for path in [output_json, log_csv]:
        if path.exists():
            path.unlink()

    env = os.environ.copy()
    env["OLLAMA_MODEL"] = model_name
    env["OLLAMA_TEMPERATURE"] = str(temperature)
    env["MAX_PASSES"] = str(max_passes)
    env["AGENT_LOG_PATH"] = str(log_csv)

    command = [
        sys.executable,
        "src/agent.py",
        str(input_json),
        str(output_json),
    ]

    print(f"Running {model_name} -> {output_json}")
    subprocess.run(command, env=env, check=True)

    return {
        "model": model_name,
        "output_json": str(output_json),
        "log_csv": str(log_csv),
    }


def write_manifest(results_dir, runs, config_path, limit):
    manifest = {
        "config": str(config_path),
        "limit": limit,
        "runs": runs,
    }
    manifest_path = results_dir / "run_manifest.json"
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return manifest_path


def summarize(results_dir):
    command = [
        sys.executable,
        "scripts/summarize_ollama_results.py",
        str(results_dir),
    ]
    subprocess.run(command, check=True)


def main():
    parser = argparse.ArgumentParser(
        description="Run the AIED 2026 agentic pipeline over configured Ollama models."
    )
    parser.add_argument(
        "config",
        nargs="?",
        default=DEFAULT_CONFIG,
        help="Path to an experiment config JSON.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Run only the first N chapters for a smoke test.",
    )
    parser.add_argument(
        "--models",
        nargs="+",
        default=None,
        help="Optional subset of model names from the config.",
    )
    args = parser.parse_args()
    config_path = Path(args.config)
    config = load_config(config_path)
    check_paths(config)
    results_dir = Path(config["results_dir"])
    input_json = Path(config["input_json"])
    temperature = config.get("temperature", 0)
    max_passes = config.get("max_passes", 3)
    models = config["models"]

    if args.models:
        requested = set(args.models)
        models = [model for model in models if model in requested]
        missing = requested - set(models)
        if missing:
            raise ValueError(f"Requested models not found in config: {', '.join(sorted(missing))}")

    results_dir.mkdir(parents=True, exist_ok=True)
    run_input = make_limited_input(input_json, results_dir, args.limit)

    runs = []
    for model in models:
        runs.append(
            run_model(
                model=model,
                input_json=run_input,
                results_dir=results_dir,
                temperature=temperature,
                max_passes=max_passes,
            )
        )

    manifest_path = write_manifest(results_dir, runs, config_path, args.limit)
    print(f"Wrote manifest: {manifest_path}")

    summarize(results_dir)
    print(f"Wrote summary: {results_dir / 'summary.csv'}")
