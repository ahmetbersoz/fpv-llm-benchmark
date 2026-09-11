"""Process LLM benchmark responses from input/results.xlsx and produce analysis.

Reads one sheet per model family (google, claude, openai, minimax, z.ai, ...)
from the input Excel, parses JSON ranking responses, computes error metrics
(Spearman, Kendall Tau, MARE, NDCG) against the ground truth, and writes a
formatted analysis workbook.

Each model sheet has the header row:
    Scenario ID | Model Version | Run | User | Date | Response

Usage:
    python -m scripts.process_results
    python -m scripts.process_results --input input/results.xlsx --output output/analysis.xlsx
"""
from __future__ import annotations

import argparse
import colorsys
import json
import math
import re
from collections import Counter, defaultdict
from pathlib import Path
from statistics import mean, pstdev
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Ground truth (Haspolat et al. 2024, Table 15)
# ---------------------------------------------------------------------------

GROUND_TRUTH = [
    "Sarıyar", "Ermenek", "Kapulukaya", "Bağıştaş", "Adıgüzel",
    "Tatar", "Kılıçkaya", "Koçköprü", "Manavgat", "Alaköprü",
    "Kemer", "Çine", "Gönen", "Arkun", "Manyas",
]

SCENARIOS = ["S1", "S2", "S3", "S4", "S5", "S6", "S7"]

# If a scenario presents the dams under alternative labels, map those labels
# back to the original names here so the ranking can be scored against
# GROUND_TRUTH. Leave empty if no relabelling is used.
S6_DAM_ALIASES: Dict[str, str] = {}

# Merge model variants into one combined model. Maps a raw "Model Version"
# value to the canonical name its runs should be counted under. Applied before
# filtering, so MODEL_FILTER refers to the merged name.
MODEL_MERGES: Dict[str, str] = {}

# Only process results for these model versions (matched against the
# "Model Version" column, after merging). Leave empty ([]) to include every model.
MODEL_FILTER: List[str] = []

# Short display names for charts (provider name omitted), keyed by the merged
# "Model Version" value. Versions not listed fall back to the raw value.
MODEL_DISPLAY_NAMES: Dict[str, str] = {}

# Optional short description of each scenario, shown in parentheses on the
# Metrics sheet column headers. Scenarios not listed are shown by ID only.
SCENARIO_DESC: Dict[str, str] = {}

# ---------------------------------------------------------------------------
# Metrics
# ---------------------------------------------------------------------------

def _rank_map(ranking: List[str]) -> Dict[str, int]:
    return {item: i + 1 for i, item in enumerate(ranking)}


def normalize_ranking(ranking: List[str], truth: List[str]) -> List[str]:
    """Return a clean 15-dam ranking suitable for scoring.

    - Drops names not present in ``truth`` (invalid / malformed LLM outputs).
    - Drops duplicates, keeping the first occurrence.
    - Appends any missing ground-truth dams to the end, preserving their
      ground-truth order, so the result is always a full permutation of
      ``truth``.
    """
    truth_set = set(truth)
    seen: set = set()
    cleaned: List[str] = []
    for item in ranking:
        name = str(item).strip()
        if name in truth_set and name not in seen:
            cleaned.append(name)
            seen.add(name)
    for name in truth:
        if name not in seen:
            cleaned.append(name)
            seen.add(name)
    return cleaned


# Graded relevance keyed by ground-truth rank (1-based), used by NDCG@5.
def _gt_relevance(truth: List[str]) -> Dict[str, int]:
    """Map each ground-truth dam to a graded relevance score by its GT rank.

    Ranks 1-3 -> 3, ranks 4-5 -> 2, ranks 6-10 -> 1, ranks 11-15 -> 0.
    """
    relevance: Dict[str, int] = {}
    for i, item in enumerate(truth):
        rank = i + 1
        if rank <= 3:
            rel = 3
        elif rank <= 5:
            rel = 2
        elif rank <= 10:
            rel = 1
        else:
            rel = 0
        relevance[item] = rel
    return relevance


def ndcg_at_k(predicted: List[str], truth: List[str], k: int = 5) -> float:
    """Graded NDCG@K using gain = 2^relevance - 1 and log2(i+2) discount.

    Relevance comes from the ground-truth rank (see ``_gt_relevance``). IDCG@K
    is computed from the ideal ordering of the same relevance values.
    """
    relevance = _gt_relevance(truth)
    dcg = sum(
        (2 ** relevance.get(predicted[i], 0) - 1) / math.log2(i + 2)
        for i in range(min(len(predicted), k))
    )
    ideal_rels = sorted(relevance.values(), reverse=True)
    idcg = sum(
        (2 ** ideal_rels[i] - 1) / math.log2(i + 2)
        for i in range(min(len(ideal_rels), k))
    )
    return dcg / idcg if idcg else 0.0


def ndcg_at_5(predicted: List[str], truth: List[str]) -> float:
    return ndcg_at_k(predicted, truth, k=5)


def spearman(predicted: List[str], truth: List[str]) -> float:
    n = len(truth)
    truth_map = _rank_map(truth)
    pred_map = _rank_map(predicted)
    d_sq = sum((pred_map.get(item, n) - truth_map[item]) ** 2 for item in truth)
    return 1 - (6 * d_sq) / (n * (n**2 - 1))


def kendall_tau(predicted: List[str], truth: List[str]) -> float:
    n = len(truth)
    truth_map = _rank_map(truth)
    pred_map = _rank_map(predicted)
    items = list(truth)
    concordant = discordant = 0
    for i in range(n):
        for j in range(i + 1, n):
            t_diff = truth_map[items[i]] - truth_map[items[j]]
            p_diff = pred_map.get(items[i], n) - pred_map.get(items[j], n)
            if t_diff * p_diff > 0:
                concordant += 1
            elif t_diff * p_diff < 0:
                discordant += 1
    pairs = n * (n - 1) / 2
    return (concordant - discordant) / pairs if pairs else 0.0


def mare(predicted: List[str], truth: List[str]) -> float:
    n = len(truth)
    truth_map = _rank_map(truth)
    pred_map = _rank_map(predicted)
    total = sum(abs(pred_map.get(item, n) - truth_map[item]) for item in truth)
    return total / n


def full_list_ndcg(predicted: List[str], truth: List[str]) -> float:
    """Legacy full-list NDCG over all 15 dams (linear relevance n-rank+1).

    Kept as a secondary metric; NDCG@5 is the primary ranking-quality metric.
    """
    n = len(truth)
    truth_map = _rank_map(truth)
    relevance = {item: n - rank + 1 for item, rank in truth_map.items()}
    dcg = sum(
        relevance.get(predicted[i], 0) / math.log2(i + 2)
        for i in range(min(len(predicted), n))
    )
    ideal_order = sorted(truth, key=lambda x: relevance[x], reverse=True)
    idcg = sum(relevance[ideal_order[i]] / math.log2(i + 2) for i in range(n))
    return dcg / idcg if idcg else 0.0


def _recall_at_k(predicted: List[str], truth: List[str], k: int) -> float:
    """Fraction of the ground-truth top K dams that appear in the predicted top K.

    Order within the top K does not matter, only set overlap. E.g. for K=3:
    ``len(set(predicted[:3]) & set(truth[:3])) / 3``.
    """
    truth_top = set(truth[:k])
    pred_top = set(predicted[:k])
    return len(truth_top & pred_top) / k if k else 0.0


def recall_at_1(predicted: List[str], truth: List[str]) -> float:
    return _recall_at_k(predicted, truth, 1)


def recall_at_3(predicted: List[str], truth: List[str]) -> float:
    return _recall_at_k(predicted, truth, 3)


def recall_at_5(predicted: List[str], truth: List[str]) -> float:
    return _recall_at_k(predicted, truth, 5)


METRIC_FUNCS = [
    ("NDCG@5", ndcg_at_5),
    ("Recall@1", recall_at_1),
    ("Recall@3", recall_at_3),
    ("Recall@5", recall_at_5),
    ("Spearman ρ", spearman),
    ("Kendall τ", kendall_tau),
    ("MARE", mare),
    ("Full-list NDCG", full_list_ndcg),
]

# Metrics where a lower value is better (all others: higher is better).
LOWER_IS_BETTER = {"MARE"}

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------

THIN = Side(border_style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
WRAP = Alignment(wrap_text=True, vertical="top")
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11, name="Arial")
SUBHEADER_FILL = PatternFill("solid", fgColor="2E75B6")
SUBHEADER_FONT = Font(bold=True, color="FFFFFF", size=10, name="Arial")
METRIC_FILL = PatternFill("solid", fgColor="D9E1F2")
GT_FILL = PatternFill("solid", fgColor="C6EFCE")
GOOD_FILL = PatternFill("solid", fgColor="C6EFCE")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
BODY_FONT = Font(size=10, name="Arial")


def _hls_palette(names: List[str]) -> Dict[str, str]:
    """Assign a stable pastel hex color to each name, evenly spread around the hue wheel."""
    n = len(names)
    palette: Dict[str, str] = {}
    for i, name in enumerate(names):
        hue = (i / n) % 1.0 if n else 0.0
        lightness = 0.78 if i % 2 == 0 else 0.86
        saturation = 0.65 if i % 2 == 0 else 0.55
        r, g, b = colorsys.hls_to_rgb(hue, lightness, saturation)
        palette[name] = f"{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}"
    return palette


def _dam_palette() -> Dict[str, str]:
    return _hls_palette(sorted(set(GROUND_TRUTH)))


# ---------------------------------------------------------------------------
# Parse input Excel
# ---------------------------------------------------------------------------

def _try_parse_json(text: str) -> Optional[Any]:
    """Parse a JSON object or array, tolerating markdown fences and surrounding text."""
    if not text or not text.strip():
        return None
    text = text.strip()
    # Strip markdown fences
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```\s*$", "", text)
    # strict=False allows raw newlines/control chars inside string values,
    # which models often emit in the "reasoning" field.
    try:
        return json.loads(text, strict=False)
    except json.JSONDecodeError:
        # Try to extract a JSON object or array from surrounding text
        match = re.search(r"(\{.*\}|\[.*\])", text, re.DOTALL)
        if match:
            try:
                return json.loads(match.group(), strict=False)
            except json.JSONDecodeError:
                pass
    return None


def _extract_ranking(data: Any) -> Tuple[List[str], str]:
    """Return (ranking, reasoning) from a parsed response.

    Accepts either a ``{"ranking": [...], "reasoning": ...}`` object or a bare
    ``[...]`` list of dam names.
    """
    if isinstance(data, list):
        return data, ""
    if isinstance(data, dict):
        return data.get("ranking", []), data.get("reasoning", "")
    return [], ""


def _salvage_ranking(text: str) -> List[str]:
    """Best-effort extraction of just the ranking array from malformed JSON.

    Used when a response is otherwise unparseable (e.g. an unclosed reasoning
    string) but contains a valid ``"ranking": [...]`` array.
    """
    match = re.search(r'"ranking"\s*:\s*(\[.*?\])', text, re.DOTALL)
    if not match:
        return []
    try:
        arr = json.loads(match.group(1), strict=False)
        return arr if isinstance(arr, list) else []
    except json.JSONDecodeError:
        return []


# Map raw sheet names to nicely formatted model-family labels.
FAMILY_LABELS = {
    "google": "Google",
    "claude": "Claude",
    "openai": "OpenAI",
    "minimax": "MiniMax",
    "z.ai": "Z.ai",
}

# Expected header (lower-cased, trimmed) identifying a model results sheet.
EXPECTED_HEADER = ["scenario id", "model version", "run", "user", "date", "response"]


def _is_results_sheet(ws) -> bool:
    header = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[1]]
    # Require the key columns to be present and in order at the start.
    return header[: len(EXPECTED_HEADER)] == EXPECTED_HEADER


def _load_results(input_path: Path) -> List[Dict[str, Any]]:
    wb = load_workbook(input_path, data_only=True)
    entries = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if not _is_results_sheet(ws):
            continue
        model_family = FAMILY_LABELS.get(sheet_name.lower(), sheet_name.capitalize())
        for row in ws.iter_rows(min_row=2, values_only=False):
            scenario = row[0].value
            model_version = row[1].value or ""
            model_version = MODEL_MERGES.get(str(model_version), model_version)
            if MODEL_FILTER and str(model_version) not in MODEL_FILTER:
                continue
            run = row[2].value or ""
            user = row[3].value or ""
            # Column 4 is Date (unused); Response is column 5.
            response_text = str(row[5].value) if row[5].value else ""

            if not scenario or not response_text.strip():
                continue

            data = _try_parse_json(response_text)
            if data is None:
                # Fall back to salvaging just the ranking array from malformed JSON.
                ranking, reasoning = _salvage_ranking(response_text), ""
                if not ranking:
                    print(f"  [WARN] {sheet_name} row {row[0].row}: could not parse JSON")
                    continue
            else:
                ranking, reasoning = _extract_ranking(data)
            if str(scenario).upper() == "S6":
                ranking = [S6_DAM_ALIASES.get(str(d).strip(), d) for d in ranking]
            if not ranking or len(ranking) < 10:
                print(f"  [WARN] {sheet_name} row {row[0].row}: ranking too short ({len(ranking)})")
                continue

            # Normalize to a clean 15-dam permutation of GROUND_TRUTH so every
            # metric receives well-formed input (drops invalid names/duplicates,
            # appends any missing dams). Applied consistently before scoring.
            ranking = normalize_ranking(ranking, GROUND_TRUTH)

            label = f"{model_family}"
            if model_version:
                label += f" {model_version}"

            entries.append({
                "model_family": model_family,
                "model_version": str(model_version),
                "label": label,
                "scenario": str(scenario).upper(),
                "run": str(run),
                "user": str(user),
                "ranking": ranking,
                "reasoning": reasoning,
            })

    entries.sort(key=lambda e: (e["model_family"], e["scenario"], e["run"]))
    return entries


# ---------------------------------------------------------------------------
# Output: Summary sheet
# ---------------------------------------------------------------------------

def _write_summary(wb: Workbook, entries: List[Dict], palette: Dict[str, str]) -> None:
    ws = wb.create_sheet("Rankings")
    n = len(GROUND_TRUTH)

    # Headers
    ws.cell(row=1, column=1, value="Rank").fill = HEADER_FILL
    ws.cell(row=1, column=1).font = HEADER_FONT
    ws.cell(row=1, column=1).alignment = CENTER
    ws.cell(row=1, column=1).border = BORDER

    ws.cell(row=1, column=2, value="Ground Truth").fill = HEADER_FILL
    ws.cell(row=1, column=2).font = HEADER_FONT
    ws.cell(row=1, column=2).alignment = CENTER
    ws.cell(row=1, column=2).border = BORDER

    for i, entry in enumerate(entries):
        col = 3 + i
        title = f"{entry['label']}\n{entry['scenario']} {entry['run']}"
        c = ws.cell(row=1, column=col, value=title)
        c.fill = HEADER_FILL
        c.font = Font(bold=True, color="FFFFFF", size=9, name="Arial")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER

    # Body
    for rank_idx in range(n):
        row = 2 + rank_idx
        ws.cell(row=row, column=1, value=rank_idx + 1).alignment = CENTER
        ws.cell(row=row, column=1).border = BORDER
        ws.cell(row=row, column=1).font = BODY_FONT

        dam = GROUND_TRUTH[rank_idx]
        c = ws.cell(row=row, column=2, value=dam)
        c.fill = GT_FILL
        c.alignment = CENTER
        c.border = BORDER
        c.font = Font(bold=True, size=10, name="Arial") if rank_idx == 0 else BODY_FONT

        for j, entry in enumerate(entries):
            col = 3 + j
            if rank_idx < len(entry["ranking"]):
                dam = entry["ranking"][rank_idx]
                cell = ws.cell(row=row, column=col, value=dam)
                cell.fill = PatternFill("solid", fgColor=palette.get(dam, "FFFFFF"))
                cell.alignment = CENTER
                cell.border = BORDER
                cell.font = BODY_FONT

    # Metrics section
    metrics_row = 2 + n + 1
    ws.cell(row=metrics_row, column=1, value="Metrics vs. Ground Truth").font = Font(bold=True, size=11, name="Arial")
    ws.cell(row=metrics_row, column=1).fill = METRIC_FILL
    metrics_row += 1

    for m_idx, (m_name, m_func) in enumerate(METRIC_FUNCS):
        row = metrics_row + m_idx
        ws.cell(row=row, column=1, value=m_name).font = Font(bold=True, name="Arial")
        ws.cell(row=row, column=1).border = BORDER
        ws.cell(row=row, column=2).border = BORDER

        for j, entry in enumerate(entries):
            col = 3 + j
            val = m_func(entry["ranking"], GROUND_TRUTH)
            cell = ws.cell(row=row, column=col, value=round(val, 4))
            cell.alignment = CENTER
            cell.border = BORDER
            cell.number_format = "0.0000"
            cell.font = BODY_FONT

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 16
    for i in range(len(entries)):
        ws.column_dimensions[get_column_letter(3 + i)].width = 16


# ---------------------------------------------------------------------------
# Output: Metrics sheet (per scenario leaderboard, best → worst)
# ---------------------------------------------------------------------------

def _write_metrics_average(
    wb: Workbook,
    entries: List[Dict],
    sheet_name: str = "Metrics",
    key: str = "label",
    group_col_label: str = "Model",
    only_version: Optional[str] = None,
) -> None:
    """Per-scenario leaderboard (best→worst), grouped by ``key`` (e.g. model label
    or contributor). Each scenario column shows Mean and population Std over the
    individual runs that fall into that group. If ``only_version`` is set, only
    runs of that model version are considered (a controlled comparison)."""
    ws = wb.create_sheet(sheet_name)

    # Group entries: group_value -> scenario -> list of rankings
    grouped: Dict[str, Dict[str, List[List[str]]]] = defaultdict(lambda: defaultdict(list))
    for e in entries:
        if only_version is not None and e["model_version"] != only_version:
            continue
        gval = e.get(key)
        if not gval:
            continue
        grouped[str(gval)][e["scenario"]].append(e["ranking"])

    models = sorted(grouped.keys())
    scenarios = [s for s in SCENARIOS if any(s in grouped[m] for m in models)]

    if not models:
        ws.cell(row=1, column=1, value="No results found.")
        return

    # Stable color per model, shared across every scenario column.
    model_palette = _hls_palette(models)

    columns = scenarios + ["Average"]
    current_row = 1

    for m_name, m_func in METRIC_FUNCS:
        reverse = m_name not in LOWER_IS_BETTER  # higher is better unless MARE

        # Section header
        c = ws.cell(row=current_row, column=1, value=m_name)
        c.font = Font(bold=True, size=12, name="Arial")
        c.fill = METRIC_FILL
        current_row += 1

        # Group header row: each column-group spans Model + Mean + Std (3 cols).
        for grp_idx, h in enumerate(columns):
            model_col = 3 * grp_idx + 1
            ws.merge_cells(start_row=current_row, start_column=model_col,
                           end_row=current_row, end_column=model_col + 2)
            for col in (model_col, model_col + 1, model_col + 2):
                cell = ws.cell(row=current_row, column=col)
                cell.fill = HEADER_FILL
                cell.border = BORDER
            desc = SCENARIO_DESC.get(h)
            title = f"{h} ({desc})" if desc else h
            c = ws.cell(row=current_row, column=model_col, value=title)
            c.font = HEADER_FONT
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        current_row += 1

        # Sub-header row: Model | Mean | Std per group.
        for grp_idx in range(len(columns)):
            model_col = 3 * grp_idx + 1
            for offset, label in ((0, group_col_label), (1, "Mean"), (2, "Std")):
                cell = ws.cell(row=current_row, column=model_col + offset, value=label)
                cell.fill = SUBHEADER_FILL
                cell.font = SUBHEADER_FONT
                cell.alignment = CENTER
                cell.border = BORDER
        current_row += 1

        # Build a best→worst sorted (model, mean, std) leaderboard for each column.
        # Std is the population std deviation across that model's individual runs.
        col_lists: List[List[Tuple[str, float, float]]] = []
        for scenario in scenarios:
            lst = []
            for model in models:
                rankings = grouped[model].get(scenario, [])
                if rankings:
                    vals = [m_func(r, GROUND_TRUTH) for r in rankings]
                    lst.append((model, mean(vals), pstdev(vals) if len(vals) > 1 else 0.0))
            lst.sort(key=lambda x: x[1], reverse=reverse)
            col_lists.append(lst)

        # Overall average column: mean of per-scenario averages; std over all runs.
        avg_list = []
        for model in models:
            per_scenario = [
                mean(m_func(r, GROUND_TRUTH) for r in grouped[model][scenario])
                for scenario in scenarios
                if grouped[model].get(scenario)
            ]
            all_runs = [
                m_func(r, GROUND_TRUTH)
                for scenario in scenarios
                for r in grouped[model].get(scenario, [])
            ]
            if per_scenario:
                avg_list.append((model, mean(per_scenario),
                                 pstdev(all_runs) if len(all_runs) > 1 else 0.0))
        avg_list.sort(key=lambda x: x[1], reverse=reverse)
        col_lists.append(avg_list)

        # Emit ranked rows: best at the top. Model (colored) | mean | std.
        max_rows = max((len(lst) for lst in col_lists), default=0)
        start = current_row
        for rank_i in range(max_rows):
            row = start + rank_i
            for grp_idx, lst in enumerate(col_lists):
                model_col = 3 * grp_idx + 1
                name_cell = ws.cell(row=row, column=model_col)
                mean_cell = ws.cell(row=row, column=model_col + 1)
                std_cell = ws.cell(row=row, column=model_col + 2)
                for cell in (name_cell, mean_cell, std_cell):
                    cell.border = BORDER
                    cell.alignment = CENTER
                    cell.font = BODY_FONT
                if rank_i < len(lst):
                    model, mean_val, std_val = lst[rank_i]
                    name_cell.value = model
                    name_cell.fill = PatternFill("solid", fgColor=model_palette.get(model, "FFFFFF"))
                    name_cell.alignment = Alignment(horizontal="left", vertical="center")
                    mean_cell.value = round(mean_val, 4)
                    mean_cell.number_format = "0.0000"
                    std_cell.value = round(std_val, 4)
                    std_cell.number_format = "0.0000"

        # Footer row: mean score across all models, per scenario / Average column.
        avg_row = start + max_rows
        for grp_idx, lst in enumerate(col_lists):
            model_col = 3 * grp_idx + 1
            name_cell = ws.cell(row=avg_row, column=model_col, value="All models avg")
            mean_cell = ws.cell(row=avg_row, column=model_col + 1)
            std_cell = ws.cell(row=avg_row, column=model_col + 2)
            for cell in (name_cell, mean_cell, std_cell):
                cell.border = BORDER
                cell.fill = METRIC_FILL
                cell.font = Font(bold=True, size=9, name="Arial")
            name_cell.alignment = Alignment(horizontal="left", vertical="center")
            mean_cell.alignment = CENTER
            std_cell.alignment = CENTER
            if lst:
                mean_cell.value = round(mean(v for _, v, _ in lst), 4)
                mean_cell.number_format = "0.0000"
        current_row = avg_row + 2  # blank row between metric tables

    # Column widths: wide model columns, narrow mean/std columns.
    for grp_idx in range(len(columns)):
        ws.column_dimensions[get_column_letter(3 * grp_idx + 1)].width = 34
        ws.column_dimensions[get_column_letter(3 * grp_idx + 2)].width = 9
        ws.column_dimensions[get_column_letter(3 * grp_idx + 3)].width = 9


# ---------------------------------------------------------------------------
# Output: Testers-by-model sheet
# ---------------------------------------------------------------------------

def _write_tester_by_model(wb: Workbook, entries: List[Dict]) -> None:
    """Compare testers on the models several of them ran, broken down per scenario.

    One block per scenario; within each, one metric table per metric whose
    columns are the shared models (run by >=2 testers in that scenario) and whose
    rows list each tester's Mean / Std / Runs for that scenario, sorted best->worst."""
    ws = wb.create_sheet("Testers by Model")

    # scenario -> tester -> model -> list of rankings
    data: Dict[str, Dict[str, Dict[str, List[List[str]]]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(list)))
    for e in entries:
        if not e["user"]:
            continue
        data[e["scenario"]][e["user"]][e["model_version"]].append(e["ranking"])

    all_testers = sorted({e["user"] for e in entries if e["user"]})
    tester_palette = _hls_palette(all_testers)
    scenarios = [s for s in SCENARIOS if s in data]

    if not scenarios:
        ws.cell(row=1, column=1, value="No tester results found.")
        return

    current_row = 1
    max_groups = 0  # widest scenario block, for setting column widths at the end

    for scenario in scenarios:
        scen_data = data[scenario]
        testers = sorted(scen_data.keys())

        # Models run by >=2 testers in this scenario.
        model_testers: Dict[str, set] = defaultdict(set)
        for tester, by_model in scen_data.items():
            for model in by_model:
                model_testers[model].add(tester)
        shared_models = sorted(m for m, ts in model_testers.items() if len(ts) >= 2)
        if not shared_models:
            continue
        max_groups = max(max_groups, len(shared_models))

        # Scenario title spanning the whole block.
        desc = SCENARIO_DESC.get(scenario)
        title = f"Scenario {scenario}" + (f" ({desc})" if desc else "")
        c = ws.cell(row=current_row, column=1, value=title)
        c.font = Font(bold=True, size=13, color="FFFFFF", name="Arial")
        c.fill = HEADER_FILL
        ws.merge_cells(start_row=current_row, start_column=1,
                       end_row=current_row, end_column=max(1, 4 * len(shared_models)))
        current_row += 1

        for m_name, m_func in METRIC_FUNCS:
            reverse = m_name not in LOWER_IS_BETTER  # higher is better unless MARE

            # Metric label row
            c = ws.cell(row=current_row, column=1, value=m_name)
            c.font = Font(bold=True, size=12, name="Arial")
            c.fill = METRIC_FILL
            current_row += 1

            # Group header row: each model spans Tester + Mean + Std + Runs (4 cols).
            for grp_idx, model in enumerate(shared_models):
                model_col = 4 * grp_idx + 1
                ws.merge_cells(start_row=current_row, start_column=model_col,
                               end_row=current_row, end_column=model_col + 3)
                for col in range(model_col, model_col + 4):
                    cell = ws.cell(row=current_row, column=col)
                    cell.fill = HEADER_FILL
                    cell.border = BORDER
                c = ws.cell(row=current_row, column=model_col, value=model)
                c.font = HEADER_FONT
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            current_row += 1

            # Sub-header row: Tester | Mean | Std | Runs per group.
            for grp_idx in range(len(shared_models)):
                model_col = 4 * grp_idx + 1
                for offset, label in ((0, "Tester"), (1, "Mean"), (2, "Std"), (3, "Runs")):
                    cell = ws.cell(row=current_row, column=model_col + offset, value=label)
                    cell.fill = SUBHEADER_FILL
                    cell.font = SUBHEADER_FONT
                    cell.alignment = CENTER
                    cell.border = BORDER
            current_row += 1

            # Best->worst (tester, mean, std, n) leaderboard per model column.
            col_lists: List[List[Tuple[str, float, float, int]]] = []
            for model in shared_models:
                lst = []
                for tester in testers:
                    rankings = scen_data[tester].get(model, [])
                    if rankings:
                        vals = [m_func(r, GROUND_TRUTH) for r in rankings]
                        lst.append((tester, mean(vals),
                                    pstdev(vals) if len(vals) > 1 else 0.0, len(vals)))
                lst.sort(key=lambda x: x[1], reverse=reverse)
                col_lists.append(lst)

            max_rows = max((len(lst) for lst in col_lists), default=0)
            start = current_row
            for rank_i in range(max_rows):
                row = start + rank_i
                for grp_idx, lst in enumerate(col_lists):
                    model_col = 4 * grp_idx + 1
                    name_cell = ws.cell(row=row, column=model_col)
                    mean_cell = ws.cell(row=row, column=model_col + 1)
                    std_cell = ws.cell(row=row, column=model_col + 2)
                    runs_cell = ws.cell(row=row, column=model_col + 3)
                    for cell in (name_cell, mean_cell, std_cell, runs_cell):
                        cell.border = BORDER
                        cell.alignment = CENTER
                        cell.font = BODY_FONT
                    if rank_i < len(lst):
                        tester, mean_val, std_val, n = lst[rank_i]
                        name_cell.value = tester
                        name_cell.fill = PatternFill("solid", fgColor=tester_palette.get(tester, "FFFFFF"))
                        name_cell.alignment = Alignment(horizontal="left", vertical="center")
                        mean_cell.value = round(mean_val, 4)
                        mean_cell.number_format = "0.0000"
                        std_cell.value = round(std_val, 4)
                        std_cell.number_format = "0.0000"
                        runs_cell.value = n

            # Footer row: mean score / total runs across all testers, per model column.
            avg_row = start + max_rows
            for grp_idx, lst in enumerate(col_lists):
                model_col = 4 * grp_idx + 1
                name_cell = ws.cell(row=avg_row, column=model_col, value="All testers avg")
                mean_cell = ws.cell(row=avg_row, column=model_col + 1)
                std_cell = ws.cell(row=avg_row, column=model_col + 2)
                runs_cell = ws.cell(row=avg_row, column=model_col + 3)
                for cell in (name_cell, mean_cell, std_cell, runs_cell):
                    cell.border = BORDER
                    cell.fill = METRIC_FILL
                    cell.font = Font(bold=True, size=9, name="Arial")
                name_cell.alignment = Alignment(horizontal="left", vertical="center")
                mean_cell.alignment = CENTER
                std_cell.alignment = CENTER
                runs_cell.alignment = CENTER
                if lst:
                    mean_cell.value = round(mean(v for _, v, _, _ in lst), 4)
                    mean_cell.number_format = "0.0000"
                    runs_cell.value = sum(n for _, _, _, n in lst)
            current_row = avg_row + 2  # blank row between metric tables

        current_row += 1  # extra gap between scenario blocks

    for grp_idx in range(max(max_groups, 1)):
        ws.column_dimensions[get_column_letter(4 * grp_idx + 1)].width = 16
        ws.column_dimensions[get_column_letter(4 * grp_idx + 2)].width = 9
        ws.column_dimensions[get_column_letter(4 * grp_idx + 3)].width = 9
        ws.column_dimensions[get_column_letter(4 * grp_idx + 4)].width = 7


# ---------------------------------------------------------------------------
# Output: charts (one grouped bar chart per metric)
# ---------------------------------------------------------------------------

def _write_charts(entries: List[Dict], outdir: Path) -> None:
    """Render one grouped bar chart per metric.

    X-axis = models (sorted by their overall average for that metric, best
    first), grouped bars = scenarios (S1, S2, ...). One PNG per metric saved
    under ``outdir``."""
    try:
        import matplotlib
        matplotlib.use("Agg")  # headless backend
        import matplotlib.pyplot as plt
    except ImportError:
        print("  [WARN] matplotlib not installed; skipping charts. "
              "Install with: pip install matplotlib")
        return

    # model version -> scenario -> list of rankings
    grouped: Dict[str, Dict[str, List[List[str]]]] = defaultdict(lambda: defaultdict(list))
    for e in entries:
        grouped[e["model_version"]][e["scenario"]].append(e["ranking"])

    models = list(grouped.keys())
    scenarios = [s for s in SCENARIOS if any(s in grouped[m] for m in models)]
    if not models or not scenarios:
        return

    outdir.mkdir(parents=True, exist_ok=True)
    cmap = plt.get_cmap("tab10")
    scen_colors = {s: cmap(i % 10) for i, s in enumerate(scenarios)}

    for m_name, m_func in METRIC_FUNCS:
        reverse = m_name not in LOWER_IS_BETTER  # higher is better unless MARE

        # Per-model per-scenario mean & population std, plus overall average for sorting.
        per_model_scen: Dict[str, Dict[str, float]] = {}
        per_model_std: Dict[str, Dict[str, float]] = {}
        overall: Dict[str, float] = {}
        for model in models:
            scen_means, scen_stds = {}, {}
            for s in scenarios:
                rs = grouped[model].get(s, [])
                if rs:
                    vals = [m_func(r, GROUND_TRUTH) for r in rs]
                    scen_means[s] = mean(vals)
                    scen_stds[s] = pstdev(vals) if len(vals) > 1 else 0.0
            if scen_means:
                per_model_scen[model] = scen_means
                per_model_std[model] = scen_stds
                overall[model] = mean(scen_means.values())

        ordered = sorted(overall, key=lambda m: overall[m], reverse=reverse)
        if not ordered:
            continue

        labels = [MODEL_DISPLAY_NAMES.get(m, m) for m in ordered]
        n_groups = len(ordered)
        n_bars = len(scenarios)
        bar_w = 0.8 / n_bars
        x = list(range(n_groups))

        fig, ax = plt.subplots(figsize=(max(8, n_groups * 1.1), 7.5))
        for b_idx, s in enumerate(scenarios):
            offsets = [i - 0.4 + bar_w * (b_idx + 0.5) for i in x]
            heights = [per_model_scen[m].get(s, 0.0) for m in ordered]
            errs = [per_model_std[m].get(s, 0.0) for m in ordered]
            ax.bar(offsets, heights, width=bar_w, label=s, color=scen_colors[s],
                   yerr=errs, capsize=2,
                   error_kw={"elinewidth": 0.8, "ecolor": "#444444", "alpha": 0.8})

        ax.set_title(f"{m_name} — by model & scenario (models sorted by average, ±1 std)",
                     fontsize=13, fontweight="bold", pad=42)
        ax.set_xlabel("Model (sorted by average)", fontsize=11, fontweight="bold")
        ax.set_ylabel(f"Mean {m_name}", fontsize=11, fontweight="bold")
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=9)
        if m_name not in LOWER_IS_BETTER:
            ax.set_ylim(0, 1.15)
        # Legend placed above the plot so it never overlaps the bars.
        ax.legend(title="Scenario", ncol=min(len(scenarios), 7), fontsize=9,
                  loc="lower center", bbox_to_anchor=(0.5, 1.02),
                  borderaxespad=0.0, frameon=False)
        ax.grid(axis="y", alpha=0.3)
        fig.tight_layout()

        safe = re.sub(r"[^A-Za-z0-9]+", "_", m_name).strip("_").lower()
        path = outdir / f"{safe}.png"
        fig.savefig(path, dpi=130)
        plt.close(fig)
        print(f"  Chart: {path}")


# ---------------------------------------------------------------------------
# Output: Reasoning sheet
# ---------------------------------------------------------------------------

def _write_reasoning(wb: Workbook, entries: List[Dict]) -> None:
    ws = wb.create_sheet("Reasoning")
    headers = ["Model", "Scenario", "Run", "User", "Reasoning"]
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = CENTER
        c.border = BORDER

    for i, e in enumerate(entries):
        row = 2 + i
        ws.cell(row=row, column=1, value=e["label"]).font = BODY_FONT
        ws.cell(row=row, column=2, value=e["scenario"]).font = BODY_FONT
        ws.cell(row=row, column=3, value=e["run"]).font = BODY_FONT
        ws.cell(row=row, column=4, value=e["user"]).font = BODY_FONT
        ws.cell(row=row, column=5, value=e["reasoning"]).font = BODY_FONT
        ws.cell(row=row, column=5).alignment = WRAP

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 100


# ---------------------------------------------------------------------------
# Output: Statistics sheet
# ---------------------------------------------------------------------------

def _style_table_header(ws, row: int, headers: List[str]) -> None:
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col_idx, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = CENTER
        c.border = BORDER


def _section_title(ws, row: int, text: str) -> None:
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(bold=True, size=12, name="Arial")
    c.fill = METRIC_FILL


def _write_statistics(wb: Workbook, entries: List[Dict]) -> None:
    """Run-count statistics: per provider/model, per user, and a provider×user matrix."""
    ws = wb.create_sheet("Statistics", 0)

    total = len(entries)
    families = sorted({e["model_family"] for e in entries})
    users = sorted({e["user"] for e in entries if e["user"]})
    models = sorted({(e["model_family"], e["model_version"]) for e in entries})
    scenarios = sorted({e["scenario"] for e in entries})

    row = 1

    # --- Overview ---------------------------------------------------------
    _section_title(ws, row, "Overview")
    row += 1
    overview = [
        ("Total runs", total),
        ("Providers", len(families)),
        ("Distinct models", len(models)),
        ("Scenarios covered", len(scenarios)),
        ("Contributors", len(users)),
    ]
    for label, val in overview:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True, name="Arial")
        ws.cell(row=row, column=1).border = BORDER
        c = ws.cell(row=row, column=2, value=val)
        c.alignment = CENTER
        c.border = BORDER
        c.font = BODY_FONT
        row += 1
    row += 1

    # --- Runs by provider -------------------------------------------------
    _section_title(ws, row, "Runs by provider")
    row += 1
    _style_table_header(ws, row, ["Provider", "Runs", "% of total"])
    row += 1
    fam_counts = Counter(e["model_family"] for e in entries)
    for fam in sorted(fam_counts, key=lambda f: -fam_counts[f]):
        ws.cell(row=row, column=1, value=fam).font = Font(bold=True, name="Arial")
        ws.cell(row=row, column=2, value=fam_counts[fam]).alignment = CENTER
        pct = ws.cell(row=row, column=3, value=fam_counts[fam] / total)
        pct.number_format = "0.0%"
        pct.alignment = CENTER
        for col in (1, 2, 3):
            ws.cell(row=row, column=col).border = BORDER
            if col > 1:
                ws.cell(row=row, column=col).font = BODY_FONT
        row += 1
    row += 1

    # --- Runs by provider + model ----------------------------------------
    _section_title(ws, row, "Runs by provider & model")
    row += 1
    _style_table_header(ws, row, ["Provider", "Model Version", "Runs", "Scenarios", "Contributors"])
    row += 1
    model_runs = defaultdict(list)
    for e in entries:
        model_runs[(e["model_family"], e["model_version"])].append(e)
    for (fam, ver) in sorted(model_runs, key=lambda k: (k[0], k[1])):
        runs = model_runs[(fam, ver)]
        scen = ", ".join(sorted({r["scenario"] for r in runs}))
        contribs = ", ".join(sorted({r["user"] for r in runs if r["user"]}))
        ws.cell(row=row, column=1, value=fam).font = BODY_FONT
        ws.cell(row=row, column=2, value=ver).font = BODY_FONT
        ws.cell(row=row, column=3, value=len(runs)).alignment = CENTER
        ws.cell(row=row, column=4, value=scen).font = BODY_FONT
        ws.cell(row=row, column=5, value=contribs).font = BODY_FONT
        for col in range(1, 6):
            ws.cell(row=row, column=col).border = BORDER
        ws.cell(row=row, column=3).font = BODY_FONT
        row += 1
    row += 1

    # --- Model × scenario matrix -----------------------------------------
    scen_cols = [s for s in SCENARIOS if s in scenarios]
    _section_title(ws, row, "Runs by model × scenario")
    row += 1
    _style_table_header(ws, row, ["Provider", "Model Version"] + scen_cols + ["Total"])
    row += 1
    ms_matrix = Counter((e["model_family"], e["model_version"], e["scenario"]) for e in entries)
    for (fam, ver) in sorted(model_runs, key=lambda k: (k[0], k[1])):
        ws.cell(row=row, column=1, value=fam).font = BODY_FONT
        ws.cell(row=row, column=2, value=ver).font = BODY_FONT
        ws.cell(row=row, column=1).border = BORDER
        ws.cell(row=row, column=2).border = BORDER
        m_total = 0
        for s_idx, scenario in enumerate(scen_cols):
            cnt = ms_matrix.get((fam, ver, scenario), 0)
            m_total += cnt
            c = ws.cell(row=row, column=3 + s_idx, value=cnt if cnt else "—")
            c.alignment = CENTER
            c.border = BORDER
            c.font = BODY_FONT
            if cnt >= 10:
                c.fill = GOOD_FILL
        c = ws.cell(row=row, column=3 + len(scen_cols), value=m_total)
        c.alignment = CENTER
        c.border = BORDER
        c.font = Font(bold=True, name="Arial")
        row += 1
    # Totals row
    ws.cell(row=row, column=1, value="Total").font = Font(bold=True, color="006100", name="Arial")
    ws.cell(row=row, column=1).fill = GOOD_FILL
    ws.cell(row=row, column=1).border = BORDER
    ws.cell(row=row, column=2).fill = GOOD_FILL
    ws.cell(row=row, column=2).border = BORDER
    scen_counts = Counter(e["scenario"] for e in entries)
    for s_idx, scenario in enumerate(scen_cols):
        c = ws.cell(row=row, column=3 + s_idx, value=scen_counts.get(scenario, 0))
        c.alignment = CENTER
        c.border = BORDER
        c.font = Font(bold=True, name="Arial")
        c.fill = GOOD_FILL
    c = ws.cell(row=row, column=3 + len(scen_cols), value=total)
    c.alignment = CENTER
    c.border = BORDER
    c.font = Font(bold=True, name="Arial")
    c.fill = GOOD_FILL
    row += 2

    # --- Runs by user -----------------------------------------------------
    _section_title(ws, row, "Runs by contributor")
    row += 1
    _style_table_header(ws, row, ["Contributor", "Runs", "% of total"])
    row += 1
    user_counts = Counter(e["user"] for e in entries if e["user"])
    for user in sorted(user_counts, key=lambda u: -user_counts[u]):
        ws.cell(row=row, column=1, value=user).font = Font(bold=True, name="Arial")
        ws.cell(row=row, column=2, value=user_counts[user]).alignment = CENTER
        pct = ws.cell(row=row, column=3, value=user_counts[user] / total)
        pct.number_format = "0.0%"
        pct.alignment = CENTER
        for col in (1, 2, 3):
            ws.cell(row=row, column=col).border = BORDER
            if col > 1:
                ws.cell(row=row, column=col).font = BODY_FONT
        row += 1
    row += 1

    # --- Provider × contributor matrix -----------------------------------
    _section_title(ws, row, "Provider × contributor (runs)")
    row += 1
    _style_table_header(ws, row, ["Provider"] + users + ["Total"])
    row += 1
    matrix = Counter((e["model_family"], e["user"]) for e in entries if e["user"])
    for fam in families:
        ws.cell(row=row, column=1, value=fam).font = Font(bold=True, name="Arial")
        ws.cell(row=row, column=1).border = BORDER
        fam_total = 0
        for u_idx, user in enumerate(users):
            cnt = matrix.get((fam, user), 0)
            fam_total += cnt
            c = ws.cell(row=row, column=2 + u_idx, value=cnt if cnt else "—")
            c.alignment = CENTER
            c.border = BORDER
            c.font = BODY_FONT
        c = ws.cell(row=row, column=2 + len(users), value=fam_total)
        c.alignment = CENTER
        c.border = BORDER
        c.font = Font(bold=True, name="Arial")
        row += 1
    # Totals row
    ws.cell(row=row, column=1, value="Total").font = Font(bold=True, color="006100", name="Arial")
    ws.cell(row=row, column=1).fill = GOOD_FILL
    ws.cell(row=row, column=1).border = BORDER
    for u_idx, user in enumerate(users):
        c = ws.cell(row=row, column=2 + u_idx, value=user_counts.get(user, 0))
        c.alignment = CENTER
        c.border = BORDER
        c.font = Font(bold=True, name="Arial")
        c.fill = GOOD_FILL
    c = ws.cell(row=row, column=2 + len(users), value=total)
    c.alignment = CENTER
    c.border = BORDER
    c.font = Font(bold=True, name="Arial")
    c.fill = GOOD_FILL

    # Column widths
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 42
    for i in range(2, 8):
        ws.column_dimensions[get_column_letter(i)].width = max(
            ws.column_dimensions[get_column_letter(i)].width or 0, 14
        )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=Path("input/results.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("output/analysis.xlsx"))
    parser.add_argument("--charts-dir", type=Path, default=Path("output/charts"),
                        help="Directory for per-metric chart PNGs.")
    args = parser.parse_args()

    if not args.input.exists():
        raise SystemExit(f"Input file not found: {args.input.resolve()}")

    print(f"Reading: {args.input.resolve()}")
    entries = _load_results(args.input)
    print(f"Found {len(entries)} valid response(s)")

    if not entries:
        raise SystemExit(
            "No valid responses found in any results_* sheet.\n"
            "Fill in Response column with JSON data, then re-run."
        )

    wb = Workbook()
    wb.remove(wb.active)

    palette = _dam_palette()
    _write_statistics(wb, entries)
    _write_summary(wb, entries, palette)
    _write_metrics_average(wb, entries)
    _write_tester_by_model(wb, entries)
    _write_reasoning(wb, entries)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.output)
    print(f"Wrote: {args.output.resolve()}")

    _write_charts(entries, args.charts_dir)


if __name__ == "__main__":
    main()
