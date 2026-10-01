# Supplementary Material S1: Consensus and Agreement Definitions

This supplement defines the agreement and consensus terms used for the human and LM-human analyses.

## Human Inter-Rater Agreement

Human inter-rater agreement is computed pairwise on overlapping annotation rows. For each overlapping item and rubric, two raters agree exactly when both assign the same integer score on the 1-5 rubric. The inter-rater agreement workbook reports exact agreement, mean absolute difference, and quadratic-weighted Cohen's kappa.

Source: `analysis/inter_rater_agreement_updated_completed_annotations.xlsx`, sheets `Agreement Summary`, `Rater Pair Summary`, and `Overlap Details`.

## Primary LM-Human Agreement

The primary LM-human agreement metric is `human_agreed_exact_accuracy` from `analysis/llm_human_agreement.xlsx`. This restricts the denominator to rubric cells where human annotators agreed on a single score. A judge score is counted as exact when it equals that agreed human score.

Source: `analysis/llm_human_agreement.xlsx`, sheets `LLM Human Summary` and `Metric Summary`.

## Broader Any-Human Agreement

The workbook also includes `exact_accuracy_any_human`. This broader diagnostic metric uses all rubric cells and counts a judge score as exact when it matches any human score assigned to the same item/rubric. For human-disagreed cells, either human score is accepted.

This broader metric is useful for sensitivity checks because it preserves human-disagreed cells instead of excluding them from the denominator.

## Modal Human Consensus for Score Distributions

For score-distribution summaries, the `Consensus Preferences` sheet in `analysis/inter_rater_agreement_updated_completed_annotations.xlsx` computes a modal human rating for each overlap item and rubric. If the human scores tie, the tied modes split one unit of weight evenly. These modal/tie-weighted values describe the distribution of human consensus preferences and are separate from the LM-human exact-agreement denominator.
