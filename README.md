# LLM Benchmark for FPV Site Selection

Companion repository for the manuscript *Evaluating Large Language Models as Decision Support Tools for Floating Photovoltaic Site Selection: A Multi-Model Benchmark* (under review).

The benchmark evaluates Large Language Models as decision support tools for Floating Photovoltaic (FPV) site selection on hydroelectric reservoirs in Türkiye, using the expert-derived ranking of Haspolat et al. (2024) as ground truth.

**Ground truth:** Haspolat et al. (2024). "Site selection of floating photovoltaic systems on hydropower reservoirs using fuzzy sine trigonometric decision-making model: Turkey as a case study." *Renewable and Sustainable Energy Reviews*, 206, 114830.

## Status

The manuscript is currently under peer review. The scenario prompts, the collected model responses, and the full analysis scripts will be released in this repository upon acceptance.

At the moment the repository contains the evaluation script that scores LLM-generated rankings against the reference ranking.

## Usage

Responses are collected in an Excel workbook (one sheet per model family) with the columns:

```
Scenario ID | Model Version | Run | User | Date | Response
```

where `Response` holds the model's JSON output (`{"ranking": [...], "reasoning": "..."}`).

```bash
pip install -r requirements.txt
python -m scripts.process_results --input input/results.xlsx --output output/analysis.xlsx
```

## Metrics

- **Spearman ρ**: Rank correlation (-1 to +1, higher = better)
- **Kendall τ**: Pairwise concordance (-1 to +1, higher = better)
- **MARE**: Mean Absolute Rank Error (lower = better)
- **NDCG**: Normalized Discounted Cumulative Gain (0 to 1, higher = better)
- **Recall@k**: Overlap of the predicted top k with the ground-truth top k (higher = better)
