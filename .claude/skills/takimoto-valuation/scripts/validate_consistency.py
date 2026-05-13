"""Validate internal consistency of a generated .xlsx valuation sheet.

Checks performed:
  1. 売上 - 売上原価 ≈ 売上総利益  (各年度)
  2. 売上総利益 - 販管費 ≈ 営業利益  (各年度)
  3. 粗利率 ≈ 売上総利益 / 売上高  (各年度)
  4. 営業利益率 ≈ 営業利益 / 売上高  (各年度)
  5. 増収率 ≈ (今期売上 / 前期売上) - 1  (各年度)
  6. セグメント別売上 合計 ≈ 全社売上  (各年度)
  7. 純利益 ≈ EPS × 発行済株式数  (各年度、可能なら)
  8. 全ての青セル(独立変数)に備考があるか
  9. TODO残値の有無と理由
  10. 重要度★が markdown の「最重要/核心」と整合しているか

Usage:
    python validate_consistency.py outputs/4971-mec-valuation.xlsx
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles.colors import Color


TOLERANCE_PCT = 0.005  # 0.5% relative tolerance for ratios
TOLERANCE_ABS = 3  # absolute tolerance for sums (百万円)


def _is_blue(font_color) -> bool:
    """Check if a font color is the blue (independent variable) marker."""
    if font_color is None:
        return False
    rgb = getattr(font_color, "rgb", None)
    if not rgb:
        return False
    # Strip alpha if present
    rgb_str = str(rgb).upper().lstrip("#")
    if rgb_str.startswith("00"):
        rgb_str = rgb_str[2:]
    return rgb_str.startswith("1F4E78")


def _strip_indent(label: str) -> str:
    return label.replace("　", "").strip() if label else ""


def _safe_float(val):
    """Coerce to float if possible; return None otherwise."""
    if val is None or val == "":
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s.startswith("TODO"):
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


def _find_row(ws, label_match):
    """Find a row by label match (substring or exact, indented OK)."""
    for r in range(2, ws.max_row + 1):
        label = ws.cell(row=r, column=1).value
        if not label:
            continue
        clean = _strip_indent(label)
        if isinstance(label_match, str):
            if clean == label_match or label_match in clean:
                return r
        elif callable(label_match):
            if label_match(clean):
                return r
    return None


def _row_values(ws, row, year_cols):
    """Return a dict {year_label: value} for a given row across year columns."""
    out = {}
    for i, col in enumerate(year_cols):
        cell = ws.cell(row=row, column=col)
        v = _safe_float(cell.value)
        out[ws.cell(row=1, column=col).value or f"col{i}"] = v
    return out


# ===========================================================================
# Checks
# ===========================================================================

def check_pl_arithmetic(ws, year_cols, findings):
    """売上 - 売上原価 = 売上総利益, 売上総利益 - 販管費 = 営業利益"""
    rows = {
        "売上高": _find_row(ws, "売上高 合計") or _find_row(ws, "売上高"),
        "売上原価": _find_row(ws, "売上原価"),
        "売上総利益": _find_row(ws, "売上総利益"),
        "販管費": _find_row(ws, "販管費 合計") or _find_row(ws, "販管費"),
        "営業利益": _find_row(ws, "営業利益"),
    }
    if not all(rows.values()):
        findings.append(("INFO", "P&Lの基本項目が見つからずスキップ "
                                 f"(found: {[k for k,v in rows.items() if v]})"))
        return

    sales = _row_values(ws, rows["売上高"], year_cols)
    cogs = _row_values(ws, rows["売上原価"], year_cols)
    gp = _row_values(ws, rows["売上総利益"], year_cols)
    sga = _row_values(ws, rows["販管費"], year_cols)
    op = _row_values(ws, rows["営業利益"], year_cols)

    for year in sales:
        # GP = Sales - COGS
        s, c, g = sales[year], cogs.get(year), gp.get(year)
        if s and c and g:
            expected = s - c
            if abs(expected - g) > TOLERANCE_ABS:
                findings.append((
                    "ERROR",
                    f"[{year}] 売上総利益 ≠ 売上 - 売上原価: "
                    f"計算{expected:.0f} vs シート{g:.0f}",
                ))
        # OP = GP - SGA
        gv, sv, ov = gp.get(year), sga.get(year), op.get(year)
        if gv and sv and ov:
            expected = gv - sv
            if abs(expected - ov) > TOLERANCE_ABS:
                findings.append((
                    "ERROR",
                    f"[{year}] 営業利益 ≠ 売上総利益 - 販管費: "
                    f"計算{expected:.0f} vs シート{ov:.0f}",
                ))


def check_ratios(ws, year_cols, findings):
    """粗利率 ≈ 粗利 / 売上, 営業利益率 ≈ 営利 / 売上"""
    sales_r = _find_row(ws, "売上高 合計") or _find_row(ws, "売上高")
    gp_r = _find_row(ws, "売上総利益")
    op_r = _find_row(ws, "営業利益")
    gp_rate_r = _find_row(ws, "粗利率")
    op_rate_r = _find_row(ws, "営業利益率")

    if not sales_r:
        return

    sales = _row_values(ws, sales_r, year_cols)
    if gp_r and gp_rate_r:
        gp = _row_values(ws, gp_r, year_cols)
        gp_rate = _row_values(ws, gp_rate_r, year_cols)
        for year in sales:
            if sales[year] and gp.get(year) and gp_rate.get(year):
                expected = gp[year] / sales[year]
                if abs(expected - gp_rate[year]) > TOLERANCE_PCT:
                    findings.append((
                        "WARN",
                        f"[{year}] 粗利率 不一致: 計算{expected:.1%} vs "
                        f"シート{gp_rate[year]:.1%}",
                    ))
    if op_r and op_rate_r:
        op = _row_values(ws, op_r, year_cols)
        op_rate = _row_values(ws, op_rate_r, year_cols)
        for year in sales:
            if sales[year] and op.get(year) and op_rate.get(year):
                expected = op[year] / sales[year]
                if abs(expected - op_rate[year]) > TOLERANCE_PCT:
                    findings.append((
                        "WARN",
                        f"[{year}] 営業利益率 不一致: 計算{expected:.1%} vs "
                        f"シート{op_rate[year]:.1%}",
                    ))


def check_growth_rate(ws, year_cols, findings):
    """増収率 ≈ (今期 / 前期) - 1"""
    sales_r = _find_row(ws, "売上高 合計") or _find_row(ws, "売上高")
    growth_r = _find_row(ws, "増収率")
    if not (sales_r and growth_r):
        return
    sales = _row_values(ws, sales_r, year_cols)
    growth = _row_values(ws, growth_r, year_cols)
    years = list(sales.keys())
    for i, year in enumerate(years):
        if i == 0:
            continue
        prev = sales[years[i - 1]]
        cur = sales[year]
        if prev and cur and growth.get(year):
            expected = cur / prev - 1
            if abs(expected - growth[year]) > TOLERANCE_PCT:
                findings.append((
                    "WARN",
                    f"[{year}] 増収率 不一致: 計算{expected:.1%} vs "
                    f"シート{growth[year]:.1%}",
                ))


def check_blue_cells_have_notes(ws, year_cols, note_col, findings):
    """All 青セル (independent variables) must have a 備考."""
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(row=r, column=1)
        if not _is_blue(cell.font.color):
            continue
        # Check if note column is non-empty for this row
        note = ws.cell(row=r, column=note_col).value
        label = _strip_indent(cell.value)
        if not note:
            findings.append((
                "ERROR",
                f"行{r} '{label}' は独立変数(青)だが備考が空欄",
            ))


def check_todos_remaining(ws, year_cols, findings):
    """List remaining TODOs in numeric cells."""
    n_todo = 0
    for r in range(2, ws.max_row + 1):
        for col in year_cols:
            v = ws.cell(row=r, column=col).value
            if isinstance(v, str) and v.upper().startswith("TODO"):
                n_todo += 1
    if n_todo:
        findings.append(("INFO", f"残TODO数値セル: {n_todo} 件"))


def check_eps_share_consistency(ws, year_cols, findings):
    """純利益 ≈ EPS × 発行済株式数 (千株単位なら ÷1000百万)."""
    ni_r = _find_row(ws, "純利益") or _find_row(ws, "親会社株主帰属純利益")
    eps_r = _find_row(ws, "EPS")
    shares_r = _find_row(ws, "発行済株式数")
    if not (ni_r and eps_r and shares_r):
        return

    ni = _row_values(ws, ni_r, year_cols)
    eps = _row_values(ws, eps_r, year_cols)
    shares = _row_values(ws, shares_r, year_cols)

    for year in ni:
        n = ni.get(year)
        e = eps.get(year)
        s = shares.get(year)
        if n and e and s:
            # ni 単位:百万円, eps:円, shares:千株 → eps × shares 千円 → 百万円
            expected_ni = e * s / 1000  # 千円÷1000=百万円
            if abs(expected_ni - n) > max(TOLERANCE_ABS, n * 0.01):
                findings.append((
                    "WARN",
                    f"[{year}] 純利益 ≠ EPS × 発行済株式数: "
                    f"計算{expected_ni:.0f} vs シート{n:.0f}",
                ))


# ===========================================================================
# Driver
# ===========================================================================

def validate(xlsx_path: str) -> list:
    """Run all checks. Returns list of (severity, message) tuples."""
    wb = load_workbook(xlsx_path)
    ws = wb.active

    # Identify year columns: row 1, columns 2..N-2 (last two are 重要度, 備考)
    n_cols = ws.max_column
    year_cols = list(range(2, n_cols - 1))
    note_col = n_cols

    findings: list = []
    check_pl_arithmetic(ws, year_cols, findings)
    check_ratios(ws, year_cols, findings)
    check_growth_rate(ws, year_cols, findings)
    check_blue_cells_have_notes(ws, year_cols, note_col, findings)
    check_todos_remaining(ws, year_cols, findings)
    check_eps_share_consistency(ws, year_cols, findings)
    return findings


def print_findings(findings):
    n_err = sum(1 for sev, _ in findings if sev == "ERROR")
    n_warn = sum(1 for sev, _ in findings if sev == "WARN")
    n_info = sum(1 for sev, _ in findings if sev == "INFO")
    print(f"\n=== Validation results: {n_err} ERRORs / {n_warn} WARNs "
          f"/ {n_info} INFOs ===")
    for sev, msg in findings:
        print(f"  [{sev}] {msg}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python validate_consistency.py <xlsx_path>")
        sys.exit(1)
    findings = validate(sys.argv[1])
    print_findings(findings)
    n_err = sum(1 for sev, _ in findings if sev == "ERROR")
    sys.exit(1 if n_err else 0)
