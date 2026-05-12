"""Reusable renderer that enforces Takimoto-Seminar spreadsheet rules.

Input: a SheetConfig describing rows, columns, charts.
Output: an .xlsx file with all formatting applied.

Rules enforced (per user spec):
- Numbers right-aligned, no decimals on integers
- Percentages displayed to 1 decimal
- 3-digit comma separators
- Indentation via row.indent (full-width spaces for the label column)
- Dependent variables (computed) shown in black; independent (assumption) in blue
- Every blue cell must have a 備考 entry (validated, prints warnings)
- Frozen first row + first two columns
- Gridlines hidden
- Charts: bar for level series, line+dot on secondary axis for ratio series
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.marker import Marker
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BLUE = "1F4E78"  # 独立変数(青)
BLACK = "000000"  # 従属変数(黒)
HEADER_GRAY = "404040"
SECTION_FILL = "EFEFEF"


@dataclass
class Row:
    """A single row in the spreadsheet."""

    label: str
    indent: int = 0  # 0 = top level, 1+ = nested factor
    values: list = field(default_factory=list)  # one entry per year column
    is_independent: bool = False  # True = blue, must have 備考
    importance: str = ""  # "", "★", "★★", "★★★"
    note: str = ""  # 備考
    number_format: str = "#,##0"  # default integer with commas
    is_section_header: bool = False  # bold, light gray fill
    is_blank: bool = False  # spacer row


@dataclass
class ChartSpec:
    """One chart definition."""

    title: str
    bar_rows: list  # row labels to plot as bars (left axis)
    line_rows: list  # row labels to plot as lines (right axis, with dots)
    anchor: str  # e.g., "K2"


@dataclass
class SheetConfig:
    """Configuration for one sheet."""

    name: str  # sheet name
    year_columns: list  # e.g., ["FY24/12 実", "FY25/12 実", "FY26/12 予", "FY27/12 予"]
    rows: list  # list[Row]
    charts: list = field(default_factory=list)  # list[ChartSpec]
    legend: str = "従属変数=黒、独立変数=青"


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

INDENT_STR = "　"  # full-width space, one per indent level


def _format_value(value, number_format: str) -> str:
    """Format a cell value according to its number_format."""
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        return value
    if number_format.endswith("%"):
        return value  # let openpyxl handle as number with format
    return value


def render(config: SheetConfig, output_path: str, also_dump_todos: bool = True) -> list:
    """Render SheetConfig to .xlsx at output_path. Returns list of TODO messages."""
    wb = Workbook()
    ws = wb.active
    ws.title = config.name
    ws.sheet_view.showGridLines = False

    todos: list = []

    # Header row 1: legend cell + year columns + 重要度 + 備考
    label_col = 1
    year_start_col = 2
    year_end_col = year_start_col + len(config.year_columns) - 1
    importance_col = year_end_col + 1
    note_col = year_end_col + 2

    ws.cell(row=1, column=label_col, value=config.legend)
    for i, year in enumerate(config.year_columns):
        ws.cell(row=1, column=year_start_col + i, value=year)
    ws.cell(row=1, column=importance_col, value="重要度")
    ws.cell(row=1, column=note_col, value="備考")

    # Header row formatting
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor=HEADER_GRAY)
    header_align = Alignment(horizontal="right", vertical="center")
    label_header_align = Alignment(horizontal="left", vertical="center")
    for col in range(1, note_col + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = label_header_align if col == 1 else header_align

    # Data rows
    row_idx = 2
    label_to_row: dict = {}  # for chart lookup
    for row_def in config.rows:
        if row_def.is_blank:
            row_idx += 1
            continue
        # Label with indentation
        label = INDENT_STR * row_def.indent + row_def.label
        label_cell = ws.cell(row=row_idx, column=label_col, value=label)

        # Determine font color
        font_color = BLUE if row_def.is_independent else BLACK
        bold = row_def.is_section_header or row_def.indent == 0
        cell_font = Font(color=font_color, bold=bold)
        label_cell.font = Font(
            color=font_color,
            bold=bold,
            size=11,
        )
        label_cell.alignment = Alignment(horizontal="left", vertical="center")

        if row_def.is_section_header:
            section_fill = PatternFill("solid", fgColor=SECTION_FILL)
            label_cell.fill = section_fill

        # Values
        for i, value in enumerate(row_def.values):
            cell = ws.cell(row=row_idx, column=year_start_col + i, value=value)
            cell.font = cell_font
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = row_def.number_format
            if row_def.is_section_header:
                cell.fill = PatternFill("solid", fgColor=SECTION_FILL)

        # Importance
        imp_cell = ws.cell(row=row_idx, column=importance_col, value=row_def.importance)
        imp_cell.font = cell_font
        imp_cell.alignment = Alignment(horizontal="right", vertical="center")

        # Note
        note_cell = ws.cell(row=row_idx, column=note_col, value=row_def.note)
        note_cell.font = cell_font
        note_cell.alignment = Alignment(horizontal="left", vertical="center")

        # Validation: blue cells must have a note
        if row_def.is_independent and not row_def.note:
            todos.append(
                f"[Sheet {config.name}] row {row_idx} '{row_def.label}' is "
                "independent (青) but has no 備考"
            )

        # TODO detection
        for i, value in enumerate(row_def.values):
            if isinstance(value, str) and value.startswith("TODO"):
                todos.append(
                    f"[Sheet {config.name}] row {row_idx} '{row_def.label}' "
                    f"col {config.year_columns[i]}: {value}"
                )

        label_to_row[row_def.label] = row_idx
        row_idx += 1

    # Column widths
    ws.column_dimensions[get_column_letter(label_col)].width = 38
    for i in range(len(config.year_columns)):
        ws.column_dimensions[get_column_letter(year_start_col + i)].width = 14
    ws.column_dimensions[get_column_letter(importance_col)].width = 8
    ws.column_dimensions[get_column_letter(note_col)].width = 50

    # Freeze panes: row 1 + columns A,B (label + first year)
    # Per spec: 行固定+列固定で「何の数字」「何年度」が常に見える
    ws.freeze_panes = "C2"

    # Charts
    for chart_spec in config.charts:
        # Build bar chart for level series (revenue, profit, etc.)
        bar_chart = BarChart()
        bar_chart.title = chart_spec.title
        bar_chart.type = "col"
        bar_chart.style = 2
        bar_chart.y_axis.title = "金額"
        bar_chart.x_axis.title = ""

        # Categories from year-row, but excluding the legend cell (col 1)
        cats = Reference(
            ws,
            min_col=year_start_col,
            max_col=year_end_col,
            min_row=1,
            max_row=1,
        )

        # Include label column in data so titles_from_data picks up series names
        for label in chart_spec.bar_rows:
            if label not in label_to_row:
                todos.append(
                    f"[Chart {chart_spec.title}] bar row '{label}' not found"
                )
                continue
            r = label_to_row[label]
            data = Reference(
                ws,
                min_col=label_col,
                max_col=year_end_col,
                min_row=r,
                max_row=r,
            )
            bar_chart.add_data(data, titles_from_data=True, from_rows=True)
        bar_chart.set_categories(cats)

        # Build line chart for ratio series on secondary axis
        if chart_spec.line_rows:
            line_chart = LineChart()
            line_chart.y_axis.crosses = "max"
            line_chart.y_axis.axId = 200
            line_chart.y_axis.title = "比率"
            for label in chart_spec.line_rows:
                if label not in label_to_row:
                    todos.append(
                        f"[Chart {chart_spec.title}] line row '{label}' not found"
                    )
                    continue
                r = label_to_row[label]
                data = Reference(
                    ws,
                    min_col=label_col,
                    max_col=year_end_col,
                    min_row=r,
                    max_row=r,
                )
                line_chart.add_data(data, titles_from_data=True, from_rows=True)
                # Dot marker on the latest series
                line_chart.series[-1].marker = Marker(symbol="circle", size=7)
                line_chart.series[-1].smooth = False
            # Overlay on the same chart frame
            bar_chart += line_chart

        bar_chart.width = 22
        bar_chart.height = 11
        ws.add_chart(bar_chart, chart_spec.anchor)

    # Save
    wb.save(output_path)

    if also_dump_todos and todos:
        print(f"\n=== TODOs in {output_path} ===")
        for t in todos:
            print(f"  - {t}")
    return todos


