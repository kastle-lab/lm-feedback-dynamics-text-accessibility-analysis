# Source

This directory contains the core rewriting pipeline and accessibility-scoring modules.

## Directory Structure

Path  | Contents
--- | ---
`agent.py`  | Runs the iterative rewrite loop and records per-iteration results.
`prompter.py`  | Sends rewrite prompts to the configured Ollama model.
`checker.py`  | Runs hard checks before a rewrite is accepted.
`judge.py`  | Combines DysText and LIX feedback into rewrite instructions.
`DysText_score.py`  | Aggregates DysText accessibility criteria into a score and feedback.
`LIX_score.py`  | Computes LIX readability and related feedback.
`failed_factual_drift.py`  | Detects likely entity or number drift between original and rewritten text.
`failed_longphrase.py`  | Detects overly long phrase spans in rewritten text.
`C*.py`  | Implements individual DysText-style accessibility criteria.
