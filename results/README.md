# Results
This directory contains generated outputs from the text-accessibility rewriting experiments, including model rewrites, per-iteration agent logs, human annotation workbooks, and legacy comparison runs.

## Directory Structure
Path  | Contents
--- | ---
`local_incremental/`  | Stores local-model outputs and agent logs for incremental feedback runs.
`local_holistic/`  | Stores local-model outputs and agent logs for holistic feedback runs.
`human_annotations/`  | Holds human annotation split workbooks used for agreement and LLM-human comparison analysis.
`legacy_gpt_results/`  | Contains earlier GPT result logs used as comparison artifacts.
`llm_judge/`  | Contains LLM-as-judge JSONL outputs and tracker CSVs for individual judge-model runs.