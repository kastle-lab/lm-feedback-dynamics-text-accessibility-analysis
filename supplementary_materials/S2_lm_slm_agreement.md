# Supplementary Material S2: LLM and SLM Agreement Means

This supplement derives the LLM and SLM agreement means under two LM-human agreement definitions: the primary human-agreed metric and the broader any-human metric.

## Grouping

The LLM evaluator group contains:

- GPT-OSS 120B
- GPT-5.2
- DeepSeek R1 70B

The SLM evaluator group contains:

- GPT-OSS 20B
- DeepSeek R1 32B
- Mistral Small 3.2
- LLaMA 3.2

## Means Using the Primary Human-Agreed Metric

The primary metric is `human_agreed_exact_accuracy`, which includes only rubric cells where the human annotators agreed on one score.

LLM mean:

`(44.1 + 42.7 + 39.8) / 3 = 42.2%`

SLM mean:

`(42.3 + 39.1 + 35.2 + 31.3) / 4 = 37.0%`

Therefore, the group means under the primary metric are approximately 42.2% for LLM evaluators and 37.0% for SLM evaluators.

## Means Using the Broader Any-Human Metric

The broader `exact_accuracy_any_human` metric in `analysis/llm_human_agreement.xlsx` includes all 1,260 rubric cells per evaluator and accepts any human rating on human-disagreed cells.

LLM mean using `exact_accuracy_any_human`:

`(47.5 + 46.0 + 42.9) / 3 = 45.5%`, rounded to approximately 46%.

SLM mean using `exact_accuracy_any_human`:

`(45.8 + 42.1 + 38.3 + 32.9) / 4 = 39.8%`, rounded to approximately 40%.

The any-human metric gives approximately 46% for LLM evaluators and 40% for SLM evaluators. These values are higher than the primary human-agreed means because the denominator is broader and human-disagreed cells are counted as exact if the LM score matches any human rating.
