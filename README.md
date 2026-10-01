# LM Feedback Dynamics Text Accessibility Analysis
This repository contains the experiment artifacts for analyzing how language-model feedback strategies affect educational text accessibility. The project rewrites textbook chapter excerpts with local and hosted language models, evaluates accessibility with rule-based checks, compares incremental and holistic feedback loops, and studies agreement between LLM-as-judge scores and human annotations. 

## Directory Structure

Path  | Contents
--- | ---
`data/`  | Contains the source chapter excerpts used as rewriting inputs.
`src/`  | Implements the rewriting agent, accessibility checks, readability scoring, feedback generation, and model prompting interface.
`scripts/`  | Contains experiment runners and analysis utilities for local model execution, result summarization, LLM judging, agreement analysis, and failure analysis.
`prompts/`  | Stores LLM-as-judge prompt templates and scoring rubrics.
`config/`  | Defines local model configuration for Ollama-based experiment runs.
`results/`  | Stores generated outputs and evaluation artifacts from the rewriting experiments.
`analysis/`  | Stores generated CSV, notebook, and workbook artifacts summarizing the experiment results.
`supplementary_materials/`  | Contains supplementary materials including consensus definitions, agreement derivations, and analysis sample sizes.
`requirements.txt`  | Lists the core Python package dependencies.

## Feedback Modes

Mode  | Description
--- | ---
Holistic  | Rewrites each chapter with a full feedback signal over the complete text.
Incremental  | Rewrites through an iterative feedback loop where failed checks produce targeted instructions for the next pass.

## LLM-as-a-Judge and Human Evaluation Dimensions

Dimension  | Description
--- | ---
Fidelity  | Measures whether the rewrite preserves propositions, relationships, qualifications, and factual claims.
Coverage  | Measures whether the rewrite retains the key topics and concepts from the original.
Structural Coherence  | Measures whether the organization and logical flow remain usable.
Readability  | Measures whether the rewritten chapter is easier and more accessible to read.

## License

This repository uses a multi-license structure.
  - Code in this repository is licensed under the Apache License 2.0. This applies to software-oriented materials such as `src/`, `scripts/`, `prompts/`, and `config/`. See [`LICENSE`](LICENSE).
  - Research materials, data, results, analysis artifacts, documentation, and README content are licensed under the Creative Commons Attribution-ShareAlike 4.0 International License. This applies to materials such as `data/`, `results/`, and `analysis/`. See [`LICENSE.docs`](LICENSE.docs).

This repository builds on materials from [`eilkou/DysText`](https://github.com/eilkou/DysText), especially the `AIED 2026` directory.
