# Supplementary Material S3: Analysis Sample Sizes

This supplement reports the sample sizes and denominators used in the human inter-rater and LM-human agreement analyses.

## Human Inter-Rater Agreement Support

The human inter-rater agreement analysis uses 49 overlapping rater-pair item comparisons. Each rubric has `N = 49` paired scores.

The pooled all-rubric summary combines the four rubric dimensions:

`49 item pairs * 4 rubric dimensions = 196 paired rubric cells`

Therefore, the pooled all-rubric summary has `N = 196`.

Source: `analysis/inter_rater_agreement_updated_completed_annotations.xlsx`, sheet `Agreement Summary`.

## LM-Human Agreement Support

The LM-human agreement analysis evaluates 315 human-annotated transformations per automated evaluator. Across four rubric dimensions, this gives:

`315 transformations * 4 rubric dimensions = 1,260 possible rubric cells per evaluator`

The primary LM-human agreement metric uses the human-agreed subset:

- Human-agreed rubric cells per evaluator: `N = 1,142`
- Human-disagreed rubric cells per evaluator excluded from the primary denominator: `N = 118`

By rubric, the human-agreed denominators are:

Rubric  | N
--- | ---
Fidelity  | 282
Coverage  | 282
Structural Coherence  | 285
Readability  | 293

Source: `analysis/llm_human_agreement.xlsx`, sheets `LLM Human Summary` and `Metric Summary`.
