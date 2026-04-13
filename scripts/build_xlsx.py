"""Build project estimation .xlsx report from decomposition JSON.

Usage:
    python scripts/build_xlsx.py input/decomposition.json output/Оценка_проекта.xlsx [--K 1.0] [--name "Project Name"]

The JSON file should contain the decomposition output with modules/tasks/phases.
Supports project_name field in JSON root, or --name CLI flag.
"""

import io
import json
import os
import sys
from datetime import date, datetime, timedelta

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def build_report_xlsx(
    project_name: str,
    modules: list[dict],
    K: float = 1.0,
) -> io.BytesIO:
    """Build xlsx report with three sheets: client view, estimation, and GANTT."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Оценка"

    # ── Styles ────────────────────────────────────────────────────────
    font_title = Font(name="Montserrat", size=18, bold=True)
    font_header = Font(name="Montserrat", size=11, bold=True)
    font_normal = Font(name="Montserrat", size=10)
    font_section = Font(name="Montserrat", size=11, bold=True)
    font_phase_title = Font(name="Montserrat", size=14, bold=True)
    font_date = Font(name="Montserrat", size=10, color="777777")

    fill_orange = PatternFill(start_color="FFA301", end_color="FFA301", fill_type="solid")
    fill_gray = PatternFill(start_color="F4F4F4", end_color="F4F4F4", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_mvp_header = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    fill_post_mvp_header = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def _apply_outer_border(ws, start_row, end_row, start_col, end_col):
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(row=r, column=c)
                top = Side("thin") if r == start_row else None
                bottom = Side("thin") if r == end_row else None
                left = Side("thin") if c == start_col else None
                right = Side("thin") if c == end_col else None
                cell.border = Border(top=top, bottom=bottom, left=left, right=right)

    # ── Column widths ─────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 35
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 12

    # ── Header ────────────────────────────────────────────────────────
    row = 2
    ws.merge_cells(f"B{row}:D{row}")
    ws.cell(row=row, column=2, value=f"Оценка проекта: {project_name}").font = font_title
    row += 1
    ws.cell(row=row, column=2, value=f"Актуально на: {date.today().strftime('%d.%m.%Y')}").font = font_date
    row += 2

    # ── Render tasks helper ───────────────────────────────────────────
    def _render_phase(ws, row, modules, phase, fill_header):
        for module in modules:
            tasks = [t for t in module.get("tasks", []) if t.get("phase", "mvp") == phase]
            if not tasks:
                continue

            ws.cell(row=row, column=2, value=module["name"]).font = font_section
            row += 1

            # Header row
            for col, val in [(2, "Специалист"), (3, "Задача"), (4, "Комментарий"), (5, "Мин, дн"), (6, "Макс, дн"), (7, "С коэф.")]:
                c = ws.cell(row=row, column=col, value=val)
                c.font = font_header
                c.fill = fill_header
                c.alignment = align_center
            _apply_outer_border(ws, row, row, 2, 7)
            row += 1

            for idx, t in enumerate(tasks):
                fill = fill_gray if idx % 2 == 0 else fill_white
                min_d = t.get("min_days", 0)
                max_d = t.get("max_days", 0)
                avg_d = (min_d + max_d) / 2
                final_d = round(avg_d * K, 1)
                for col, val in [(2, t["specialist"]), (3, t["task"]), (4, t.get("comment", "")), (5, min_d), (6, max_d), (7, final_d)]:
                    c = ws.cell(row=row, column=col, value=val)
                    c.font = font_normal
                    c.fill = fill
                    c.alignment = align_center if col >= 5 else align_left
                ws.cell(row=row, column=2).border = Border(left=Side("thin"))
                ws.cell(row=row, column=7).border = Border(right=Side("thin"))
                row += 1

            # Bottom border
            if tasks:
                for c in range(2, 8):
                    cell = ws.cell(row=row - 1, column=c)
                    existing = cell.border
                    cell.border = Border(left=existing.left, right=existing.right, top=existing.top, bottom=Side("thin"))
            row += 1
        return row

    # ── MVP ───────────────────────────────────────────────────────────
    ws.merge_cells(f"B{row}:G{row}")
    ws.cell(row=row, column=2, value="MVP — Базовый скоуп").font = font_phase_title
    ws.cell(row=row, column=2).fill = fill_mvp_header
    row += 1
    row = _render_phase(ws, row, modules, "mvp", fill_orange)

    # MVP subtotal
    mvp_min = sum(t.get("min_days", 0) for m in modules for t in m.get("tasks", []) if t.get("phase", "mvp") == "mvp")
    mvp_max = sum(t.get("max_days", 0) for m in modules for t in m.get("tasks", []) if t.get("phase", "mvp") == "mvp")
    ws.cell(row=row, column=3, value="Итого MVP").font = Font(name="Montserrat", bold=True)
    ws.cell(row=row, column=5, value=mvp_min).font = Font(name="Montserrat", bold=True)
    ws.cell(row=row, column=5).alignment = align_center
    ws.cell(row=row, column=6, value=mvp_max).font = Font(name="Montserrat", bold=True)
    ws.cell(row=row, column=6).alignment = align_center
    ws.cell(row=row, column=7, value=round((mvp_min + mvp_max) / 2 * K, 1)).font = Font(name="Montserrat", bold=True)
    ws.cell(row=row, column=7).alignment = align_center
    row += 2

    # ── Post-MVP ──────────────────────────────────────────────────────
    has_post_mvp = any(t.get("phase") == "post-mvp" for m in modules for t in m.get("tasks", []))
    if has_post_mvp:
        ws.merge_cells(f"B{row}:G{row}")
        ws.cell(row=row, column=2, value="Post-MVP — Расширения и усложнения").font = font_phase_title
        ws.cell(row=row, column=2).fill = fill_post_mvp_header
        row += 1
        row = _render_phase(ws, row, modules, "post-mvp", PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid"))

        pm_min = sum(t.get("min_days", 0) for m in modules for t in m.get("tasks", []) if t.get("phase") == "post-mvp")
        pm_max = sum(t.get("max_days", 0) for m in modules for t in m.get("tasks", []) if t.get("phase") == "post-mvp")
        ws.cell(row=row, column=3, value="Итого Post-MVP").font = Font(name="Montserrat", bold=True)
        ws.cell(row=row, column=5, value=pm_min).font = Font(name="Montserrat", bold=True)
        ws.cell(row=row, column=5).alignment = align_center
        ws.cell(row=row, column=6, value=pm_max).font = Font(name="Montserrat", bold=True)
        ws.cell(row=row, column=6).alignment = align_center
        ws.cell(row=row, column=7, value=round((pm_min + pm_max) / 2 * K, 1)).font = Font(name="Montserrat", bold=True)
        ws.cell(row=row, column=7).alignment = align_center
        row += 2

    # ── Grand total ───────────────────────────────────────────────────
    all_min = sum(t.get("min_days", 0) for m in modules for t in m.get("tasks", []))
    all_max = sum(t.get("max_days", 0) for m in modules for t in m.get("tasks", []))
    ws.cell(row=row, column=3, value="ИТОГО").font = Font(name="Montserrat", size=12, bold=True)
    ws.cell(row=row, column=5, value=all_min).font = Font(name="Montserrat", size=12, bold=True)
    ws.cell(row=row, column=5).alignment = align_center
    ws.cell(row=row, column=6, value=all_max).font = Font(name="Montserrat", size=12, bold=True)
    ws.cell(row=row, column=6).alignment = align_center
    ws.cell(row=row, column=7, value=round((all_min + all_max) / 2 * K, 1)).font = Font(name="Montserrat", size=12, bold=True)
    ws.cell(row=row, column=7).alignment = align_center

    # ── Sheet 2: GANTT ────────────────────────────────────────────────
    _build_gantt_sheet(wb, modules, K)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _build_gantt_sheet(wb: Workbook, modules: list[dict], K: float):
    """Add GANTT Chart sheet."""

    def next_workday(dt):
        while dt.weekday() >= 5:
            dt += timedelta(days=1)
        return dt

    def add_workdays(start_dt, num_workdays):
        cur = next_workday(start_dt)
        counted = 1
        while counted < num_workdays:
            cur += timedelta(days=1)
            if cur.weekday() < 5:
                counted += 1
        return cur

    PHASE_COLORS = [
        {"header": "1F4E79", "fill": "D6E4F0", "bar": "5B9BD5"},
        {"header": "7B2D26", "fill": "F2DCDB", "bar": "C0504D"},
        {"header": "4F6228", "fill": "EBF1DE", "bar": "9BBB59"},
        {"header": "31859C", "fill": "DAEEF3", "bar": "4BACC6"},
        {"header": "E36C09", "fill": "FDE9D9", "bar": "F79646"},
        {"header": "60497A", "fill": "E4DFEC", "bar": "8064A2"},
        {"header": "4A452A", "fill": "F2F2E6", "bar": "948A54"},
    ]

    month_names_ru = {
        1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
        5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
        9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
    }

    DATA_COL_START = 7

    specialist_tasks: dict[str, list[dict]] = {}
    for module in modules:
        for t in module.get("tasks", []):
            spec = t["specialist"]
            if spec not in specialist_tasks:
                specialist_tasks[spec] = []
            avg_d = (t.get("min_days", 0) + t.get("max_days", 0)) / 2
            specialist_tasks[spec].append({"task": t["task"], "module": module["name"], "duration": max(1, round(avg_d * K))})

    if not specialist_tasks:
        return

    spec_list = list(specialist_tasks.keys())
    spec_colors = {spec: PHASE_COLORS[i % len(PHASE_COLORS)] for i, spec in enumerate(spec_list)}

    project_start_raw = next_workday(date.today())
    project_start_dt = datetime(project_start_raw.year, project_start_raw.month, project_start_raw.day)

    scheduled_tasks = []
    for spec in spec_list:
        current_start = project_start_dt
        for t in specialist_tasks[spec]:
            start = next_workday(current_start)
            end = add_workdays(start, t["duration"])
            scheduled_tasks.append((spec, t["module"], t["task"], start, end, t["duration"]))
            current_start = end + timedelta(days=1)

    if not scheduled_tasks:
        return

    project_start = min(t[3] for t in scheduled_tasks)
    while project_start.weekday() != 0:
        project_start -= timedelta(days=1)
    project_end = max(t[4] for t in scheduled_tasks)
    while project_end.weekday() != 4:
        project_end += timedelta(days=1)

    all_days = []
    d = project_start
    while d <= project_end:
        all_days.append(d)
        d += timedelta(days=1)
    num_days = len(all_days)

    ws = wb.create_sheet("GANTT Chart")

    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    weekend_header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
    date_label_fill = PatternFill(start_color="D6DCE4", end_color="D6DCE4", fill_type="solid")
    thin_border = Border(left=Side("thin", color="BFBFBF"), right=Side("thin", color="BFBFBF"), top=Side("thin", color="BFBFBF"), bottom=Side("thin", color="BFBFBF"))
    week_sep_border = Border(left=Side("thin", color="BFBFBF"), right=Side("medium", color="808080"), top=Side("thin", color="BFBFBF"), bottom=Side("thin", color="BFBFBF"))

    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    month_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    task_font = Font(name="Calibri", size=9)
    phase_header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    center_align = Alignment(horizontal="center", vertical="center")
    left_wrap = Alignment(vertical="center", wrap_text=True)
    weekday_names = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 42
    ws.column_dimensions["D"].width = 11
    ws.column_dimensions["E"].width = 6
    ws.column_dimensions["F"].width = 11
    for i in range(num_days):
        col_letter = get_column_letter(DATA_COL_START + i)
        ws.column_dimensions[col_letter].width = 2.5 if all_days[i].weekday() >= 5 else 3.8

    def _border_for_day(day):
        return week_sep_border if day.weekday() == 6 else thin_border

    # Row 1: Months
    ws.row_dimensions[1].height = 20
    for c in range(1, 7):
        ws.cell(row=1, column=c).fill = header_fill
    month_spans = []
    if all_days:
        cur_month = (all_days[0].year, all_days[0].month)
        span_start = 0
        for i, day in enumerate(all_days):
            m = (day.year, day.month)
            if m != cur_month:
                month_spans.append((cur_month, span_start, i - 1))
                cur_month = m
                span_start = i
        month_spans.append((cur_month, span_start, len(all_days) - 1))

    for (year, month), start_idx, end_idx in month_spans:
        start_col = DATA_COL_START + start_idx
        end_col = DATA_COL_START + end_idx
        if end_col > start_col:
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
        cell = ws.cell(row=1, column=start_col, value=f"{month_names_ru[month]} {year}")
        cell.font = month_font
        cell.fill = header_fill
        cell.alignment = center_align

    # Row 2: Headers + day numbers
    ws.row_dimensions[2].height = 18
    for i, val in enumerate(["Фаза", "Роль", "Задача", "Старт", "Дней", "Конец"]):
        cell = ws.cell(row=2, column=i + 1, value=val)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
    for i, day in enumerate(all_days):
        cell = ws.cell(row=2, column=DATA_COL_START + i, value=day.day)
        cell.font = Font(name="Calibri", size=7, color="44546A")
        cell.fill = weekend_header_fill if day.weekday() >= 5 else date_label_fill
        cell.alignment = center_align

    # Row 3: Weekday names
    ws.row_dimensions[3].height = 14
    for i, day in enumerate(all_days):
        cell = ws.cell(row=3, column=DATA_COL_START + i, value=weekday_names[day.weekday()])
        cell.font = Font(name="Calibri", size=6, color="808080")
        cell.fill = weekend_header_fill if day.weekday() >= 5 else date_label_fill
        cell.alignment = center_align

    # Data rows
    row = 4
    current_spec = None
    for spec, module_name, task_name, start_dt, end_dt, duration in scheduled_tasks:
        colors = spec_colors[spec]
        fill_phase = PatternFill(start_color=colors["fill"], end_color=colors["fill"], fill_type="solid")
        fill_bar = PatternFill(start_color=colors["bar"], end_color=colors["bar"], fill_type="solid")
        fill_header_spec = PatternFill(start_color=colors["header"], end_color=colors["header"], fill_type="solid")

        if spec != current_spec:
            current_spec = spec
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            cell = ws.cell(row=row, column=1, value=spec)
            cell.font = phase_header_font
            cell.fill = fill_header_spec
            for c in range(1, 7):
                ws.cell(row=row, column=c).fill = fill_header_spec
            for i, day in enumerate(all_days):
                ws.cell(row=row, column=DATA_COL_START + i).fill = fill_header_spec
            row += 1

        ws.row_dimensions[row].height = 20
        for col, val, align in [(1, module_name, left_wrap), (2, spec, left_wrap), (3, task_name, left_wrap), (4, start_dt.strftime("%d.%m.%y"), center_align), (5, duration, center_align), (6, end_dt.strftime("%d.%m.%y"), center_align)]:
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = task_font
            cell.fill = fill_phase
            cell.alignment = align

        for i, day in enumerate(all_days):
            cell = ws.cell(row=row, column=DATA_COL_START + i)
            if start_dt <= day <= end_dt and day.weekday() < 5:
                cell.fill = fill_bar
            else:
                cell.fill = fill_phase

        row += 1

    ws.freeze_panes = "G4"


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python scripts/build_xlsx.py <decomposition.json> <output.xlsx> [--K 1.0]")
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8") as f:
        data = json.load(f)

    K = 1.0
    if "--K" in sys.argv:
        K = float(sys.argv[sys.argv.index("--K") + 1])

    modules = data.get("modules", data) if isinstance(data, dict) else data

    project_name = "Проект"
    if "--name" in sys.argv:
        project_name = sys.argv[sys.argv.index("--name") + 1]
    elif isinstance(data, dict) and "project_name" in data:
        project_name = data["project_name"]

    result = build_report_xlsx(project_name, modules, K)

    with open(sys.argv[2], "wb") as f:
        f.write(result.read())

    print(f"Saved: {sys.argv[2]}")
