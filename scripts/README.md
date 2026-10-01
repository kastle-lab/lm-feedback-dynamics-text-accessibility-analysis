# Scripts
This directory contains experiment runners and analysis utilities for the text-accessibility rewriting workflow.

## Directory Structure
Path  | Contents
--- | ---
`run_ollama_models.py`  | Runs configured Ollama models over the chapter input file.
`ollama_interface.py`  | Provides shared Ollama client helpers and model-call behavior.
`summarize_ollama_results.py`  | Summarizes per-model CSV logs from local experiment runs.
`run_llm_judge.py`  | Runs an LLM-as-judge prompt over generated rewrite results.
`consolidate_llm_judge_results.py`  | Consolidates judge JSONL outputs by source row.
`analyze_inter_rater_agreement.py`  | Computes agreement statistics across human annotation workbooks.
`analyze_llm_human_agreement.py`  | Compares LLM judge scores against human annotation scores.
`generate_results_summary_workbook.py`  | Generates a workbook summarizing experiment results.
`generate_deep_failure_analysis.py`  | Generates deeper failure-analysis summaries from agent logs.
