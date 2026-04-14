"""Build project estimation .xlsx report from decomposition JSON.

Usage:
    python scripts/build_xlsx.py input/decomposition.json output/Оценка_проекта.xlsx [--K 1.0] [--name "Project Name"]

Generates 3 sheets: "Для клиента", "Оценка", "GANTT Chart".
Matches the original template design from the Telegram bot version.
"""

import io
import json
import os
import sys
from datetime import date, datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def build_report_xlsx(
    project_name: str,
    modules: list[dict],
    K: float = 1.0,
) -> io.BytesIO:
    """Build client-facing xlsx report with 3 sheets."""
    wb = Workbook()

    # ── Sheet 1: Для клиента ─────────────────────────────────────────
    _build_client_sheet(wb, project_name, modules, K)

    # ── Sheet 2: Оценка ──────────────────────────────────────────────
    _build_estimation_sheet(wb, modules, K)

    # ── Sheet 3: GANTT Chart ─────────────────────────────────────────
    _build_gantt_sheet(wb, modules, K)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════
# Sheet 1: Для клиента
# ═══════════════════════════════════════════════════════════════════════

def _build_client_sheet(wb: Workbook, project_name: str, modules: list[dict], K: float):
    ws = wb.active
    ws.title = "Для клиента"

    font_title = Font(name="Montserrat", size=24)
    font_subtitle = Font(name="Montserrat", size=10)
    font_date = Font(name="Montserrat", size=10, color="777777")
    font_header = Font(name="Montserrat", size=12)
    font_normal = Font(name="Montserrat", size=10)
    font_section = Font(name="Montserrat", size=11, bold=True)

    fill_orange = PatternFill(start_color="FFA301", end_color="FFA301", fill_type="solid")
    fill_gray = PatternFill(start_color="F4F4F4", end_color="F4F4F4", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    def _apply_outer_border(ws, start_row, end_row, start_col, end_col):
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(row=r, column=c)
                top = Side("thin") if r == start_row else None
                bottom = Side("thin") if r == end_row else None
                left = Side("thin") if c == start_col else None
                right = Side("thin") if c == end_col else None
                cell.border = Border(top=top, bottom=bottom, left=left, right=right)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 7.8
    ws.column_dimensions["B"].width = 21
    ws.column_dimensions["C"].width = 39.5
    ws.column_dimensions["D"].width = 35
    ws.column_dimensions["E"].width = 27

    # Logo
    logo_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "metalamp-logo.png")
    if os.path.exists(logo_path):
        from openpyxl.drawing.image import Image as XlImage
        img = XlImage(logo_path)
        orig_w, orig_h = img.width, img.height
        target_w = 300
        img.width = target_w
        img.height = int(orig_h * (target_w / max(orig_w, 1)))
        ws.add_image(img, "B2")

    # Title
    row = 5
    ws.row_dimensions[row].height = 38.25
    ws.merge_cells(f"B{row}:C{row}")
    ws.cell(row=row, column=2, value="Оценка проекта ").font = font_title
    ws.merge_cells(f"D{row}:E{row}")
    ws.cell(row=row, column=4, value=project_name).font = font_title

    row = 6
    ws.row_dimensions[row].height = 38.25
    ws.merge_cells(f"B{row}:E{row}")
    ws.cell(row=row, column=2, value="В стоимость входит тестирование, работа менеджера").font = font_subtitle

    row = 7
    ws.row_dimensions[row].height = 31.5
    ws.merge_cells(f"B{row}:E{row}")
    ws.cell(row=row, column=2, value="В течение 3 месяцев мы бесплатно устраняем технические ошибки (техподдержка)").font = font_subtitle

    row = 8
    ws.cell(row=row, column=2, value=f"Актуально на: {date.today().strftime('%d.%m.%Y')}").font = font_date

    # Summary: compute total hours and specialists
    spec_hours: dict[str, float] = {}
    for m in modules:
        for t in m.get("tasks", []):
            name = t["specialist"]
            avg = (t.get("min_days", 0) + t.get("max_days", 0)) / 2
            hours = avg * K * 8
            spec_hours[name] = spec_hours.get(name, 0) + hours

    total_hours = sum(spec_hours.values())
    team_size = len(spec_hours)

    row = 9
    ws.row_dimensions[row].height = 30
    for col, val in [(2, "Команда проекта,\nчеловек"), (3, "Длительность проекта,\nчасы")]:
        c = ws.cell(row=row, column=col, value=val)
        c.font = font_header
        c.fill = fill_orange
        c.alignment = align_center
    _apply_outer_border(ws, 9, 9, 2, 3)

    row = 10
    ws.row_dimensions[row].height = 15
    for col, val in [(2, team_size), (3, round(total_hours))]:
        c = ws.cell(row=row, column=col, value=val)
        c.font = font_normal
        c.alignment = align_center
    _apply_outer_border(ws, 10, 10, 2, 3)

    # Specialists table
    row = 12
    ws.row_dimensions[row].height = 15
    for col, val in [(2, "Специалисты"), (3, "Занятость, недели"), (4, "Занятость, часы")]:
        c = ws.cell(row=row, column=col, value=val)
        c.font = font_header
        c.fill = fill_orange
        c.alignment = align_center
    _apply_outer_border(ws, 12, 12, 2, 4)

    spec_start_row = row + 1
    for idx, (name, hours) in enumerate(spec_hours.items()):
        row = spec_start_row + idx
        weeks = round(hours / 40, 1)
        fill = fill_gray if idx % 2 == 0 else fill_white
        for col, val in [(2, name), (3, weeks), (4, round(hours))]:
            c = ws.cell(row=row, column=col, value=val)
            c.font = font_normal
            c.fill = fill
            c.alignment = align_center if col >= 3 else align_left
        ws.cell(row=row, column=2).border = Border(left=Side("thin"))
        ws.cell(row=row, column=4).border = Border(right=Side("thin"))

    last_spec_row = spec_start_row + len(spec_hours) - 1
    for c in range(2, 5):
        cell = ws.cell(row=last_spec_row, column=c)
        existing = cell.border
        cell.border = Border(left=existing.left, right=existing.right, top=existing.top, bottom=Side("thin"))

    row = last_spec_row + 3

    # Task decomposition by modules
    for module in modules:
        ws.row_dimensions[row].height = 13.8
        ws.cell(row=row, column=2, value=module["name"]).font = font_section
        row += 1

        ws.row_dimensions[row].height = 15
        for col, val in [(2, "Специалист"), (3, "Задача"), (4, "Комментарий"), (5, "Оценка, дни")]:
            c = ws.cell(row=row, column=col, value=val)
            c.font = font_header
            c.fill = fill_orange
            c.alignment = align_center
        _apply_outer_border(ws, row, row, 2, 5)
        row += 1

        for idx, t in enumerate(module.get("tasks", [])):
            fill = fill_gray if idx % 2 == 0 else fill_white
            avg_d = (t["min_days"] + t["max_days"]) / 2
            days_str = str(round(avg_d * K, 1))
            phase_tag = " [Post-MVP]" if t.get("phase") == "post-mvp" else ""
            for col, val in [(2, t["specialist"]), (3, t["task"] + phase_tag), (4, t.get("comment", "")), (5, days_str)]:
                c = ws.cell(row=row, column=col, value=val)
                c.font = font_normal
                c.fill = fill
                c.alignment = align_center if col == 5 else align_left
            ws.cell(row=row, column=2).border = Border(left=Side("thin"))
            ws.cell(row=row, column=5).border = Border(right=Side("thin"))
            row += 1

        if module.get("tasks"):
            last_task_row = row - 1
            for c in range(2, 6):
                cell = ws.cell(row=last_task_row, column=c)
                existing = cell.border
                cell.border = Border(left=existing.left, right=existing.right, top=existing.top, bottom=Side("thin"))

        row += 1


# ═══════════════════════════════════════════════════════════════════════
# Sheet 2: Оценка
# ═══════════════════════════════════════════════════════════════════════

def _build_estimation_sheet(wb: Workbook, modules: list[dict], K: float):
    ws = wb.create_sheet("Оценка")

    font_bold = Font(name="Arial", bold=True)
    font_normal = Font(name="Arial")
    font_header = Font(name="Arial", bold=True)
    font_module = Font(name="Arial", size=11, bold=True)
    font_task = Font(name="Arial", size=11)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 19
    ws.column_dimensions["C"].width = 25
    ws.column_dimensions["D"].width = 29
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 13
    ws.column_dimensions["G"].width = 13
    ws.column_dimensions["H"].width = 15

    # Summary headers
    row = 4
    for col, val in [(2, "Дни минимум"), (3, "Дни максимум"), (4, f"Недель с коф. (K={K:.2f})")]:
        c = ws.cell(row=row, column=col, value=val)
        c.font = font_bold
        c.alignment = align_center if col >= 3 else align_left

    # Summary per specialist
    spec_summary: dict[str, dict] = {}
    for m in modules:
        for t in m.get("tasks", []):
            name = t["specialist"]
            if name not in spec_summary:
                spec_summary[name] = {"min": 0, "max": 0}
            spec_summary[name]["min"] += t.get("min_days", 0)
            spec_summary[name]["max"] += t.get("max_days", 0)

    summary_start = 5
    spec_number_map = {}
    for idx, (name, d) in enumerate(spec_summary.items()):
        r = summary_start + idx
        spec_num = idx + 1
        spec_number_map[name] = spec_num
        avg = (d["min"] + d["max"]) / 2
        weeks_k = round(avg * K / 5, 2)
        ws.cell(row=r, column=1, value=spec_num).font = font_normal
        ws.cell(row=r, column=1).alignment = align_center
        ws.cell(row=r, column=2, value=d["min"]).font = font_normal
        ws.cell(row=r, column=2).alignment = align_right
        ws.cell(row=r, column=3, value=d["max"]).font = font_normal
        ws.cell(row=r, column=3).alignment = align_right
        ws.cell(row=r, column=4, value=weeks_k).font = font_bold
        ws.cell(row=r, column=4).alignment = align_right
        ws.cell(row=r, column=5, value=name).font = font_normal

    # Header row for task table
    row = summary_start + len(spec_summary) + 1
    headers = {1: "Распределение работ", 2: "Вид работ", 3: "Задача", 4: "Комментарий", 6: "Мин. дни", 7: "Макс. дни", 8: "Итого с коэф."}
    for col, val in headers.items():
        c = ws.cell(row=row, column=col, value=val)
        c.font = font_header
        c.alignment = align_center
    row += 1

    # Module headers and task rows
    for module in modules:
        ws.merge_cells(f"C{row}:D{row}")
        c = ws.cell(row=row, column=3, value=module["name"])
        c.font = font_module
        c.alignment = align_center
        row += 1

        for t in module.get("tasks", []):
            spec_name = t["specialist"]
            spec_n = spec_number_map.get(spec_name, "")
            min_d = t.get("min_days", 0)
            max_d = t.get("max_days", 0)
            avg_d = (min_d + max_d) / 2
            final_d = round(avg_d * K, 1)

            ws.cell(row=row, column=1, value=spec_n).font = font_normal
            ws.cell(row=row, column=1).alignment = align_right
            ws.cell(row=row, column=2, value=spec_name).font = font_bold
            ws.cell(row=row, column=2).alignment = align_left
            ws.cell(row=row, column=3, value=t["task"]).font = font_task
            ws.cell(row=row, column=3).alignment = align_left
            ws.cell(row=row, column=4, value=t.get("comment", "")).font = font_task
            ws.cell(row=row, column=4).alignment = align_left
            ws.cell(row=row, column=6, value=min_d).font = font_normal
            ws.cell(row=row, column=6).alignment = align_center
            ws.cell(row=row, column=7, value=max_d).font = font_normal
            ws.cell(row=row, column=7).alignment = align_center
            ws.cell(row=row, column=8, value=final_d).font = font_normal
            ws.cell(row=row, column=8).alignment = align_center
            row += 1


# ═══════════════════════════════════════════════════════════════════════
# Sheet 3: GANTT Chart (exact copy from original bot.py)
# ═══════════════════════════════════════════════════════════════════════

def _build_gantt_sheet(wb: Workbook, modules: list[dict], K: float):
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

    # Build flat task list grouped by specialist
    specialist_tasks: dict[str, list[dict]] = {}
    for module in modules:
        for t in module.get("tasks", []):
            spec = t["specialist"]
            if spec not in specialist_tasks:
                specialist_tasks[spec] = []
            min_d = t.get("min_days", 0)
            max_d = t.get("max_days", 0)
            avg_d = (min_d + max_d) / 2
            duration_days = max(1, round(avg_d * K))
            specialist_tasks[spec].append({"task": t["task"], "duration": duration_days})

    if not specialist_tasks:
        return

    spec_list = list(specialist_tasks.keys())
    spec_colors = {}
    for i, spec in enumerate(spec_list):
        spec_colors[spec] = PHASE_COLORS[i % len(PHASE_COLORS)]

    project_start_raw = next_workday(date.today())
    project_start_dt = datetime(project_start_raw.year, project_start_raw.month, project_start_raw.day)

    # Schedule: (specialist, task_name, start_dt, end_dt, duration)
    scheduled_tasks = []
    for spec in spec_list:
        current_start = project_start_dt
        for t in specialist_tasks[spec]:
            start = next_workday(current_start)
            end = add_workdays(start, t["duration"])
            scheduled_tasks.append((spec, t["task"], start, end, t["duration"]))
            current_start = end + timedelta(days=1)

    if not scheduled_tasks:
        return

    project_start = min(t[2] for t in scheduled_tasks)
    while project_start.weekday() != 0:
        project_start -= timedelta(days=1)
    project_end = max(t[3] for t in scheduled_tasks)
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
    day_num_font = Font(name="Calibri", size=7, color="44546A")
    weekday_font = Font(name="Calibri", size=6, color="808080")
    weekday_bold_font = Font(name="Calibri", size=6, bold=True, color="999999")
    task_font = Font(name="Calibri", size=9)
    phase_header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    center_align = Alignment(horizontal="center", vertical="center")
    left_wrap = Alignment(vertical="center", wrap_text=True)
    phase_align = Alignment(horizontal="left", vertical="center")

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
        ws.cell(row=1, column=c).border = thin_border

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
        for ci in range(start_col, end_col + 1):
            c = ws.cell(row=1, column=ci)
            c.fill = header_fill
            c.border = _border_for_day(all_days[ci - DATA_COL_START])

    # Row 2: Headers + day numbers
    ws.row_dimensions[2].height = 18
    for i, val in enumerate(["Фаза", "Роль", "Задача", "Старт", "Дней", "Конец"]):
        cell = ws.cell(row=2, column=i + 1, value=val)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

    for i, day in enumerate(all_days):
        col = DATA_COL_START + i
        cell = ws.cell(row=2, column=col, value=day.day)
        cell.font = day_num_font
        cell.fill = weekend_header_fill if day.weekday() >= 5 else date_label_fill
        cell.alignment = center_align
        cell.border = _border_for_day(day)

    # Row 3: Weekday names
    ws.row_dimensions[3].height = 14
    for c in range(1, 7):
        ws.cell(row=3, column=c).fill = date_label_fill
        ws.cell(row=3, column=c).border = thin_border

    for i, day in enumerate(all_days):
        col = DATA_COL_START + i
        cell = ws.cell(row=3, column=col, value=weekday_names[day.weekday()])
        cell.font = weekday_bold_font if day.weekday() >= 5 else weekday_font
        cell.fill = weekend_header_fill if day.weekday() >= 5 else date_label_fill
        cell.alignment = center_align
        cell.border = _border_for_day(day)

    # Data rows
    row = 4
    current_spec = None

    for spec, task_name, start_dt, end_dt, duration in scheduled_tasks:
        colors = spec_colors[spec]

        rb = int(colors["fill"][:2], 16)
        gb = int(colors["fill"][2:4], 16)
        bb = int(colors["fill"][4:6], 16)
        wknd_hex = f"{max(0, rb - 25):02X}{max(0, gb - 25):02X}{max(0, bb - 25):02X}"

        fill_phase = PatternFill(start_color=colors["fill"], end_color=colors["fill"], fill_type="solid")
        fill_bar = PatternFill(start_color=colors["bar"], end_color=colors["bar"], fill_type="solid")
        fill_wknd = PatternFill(start_color=wknd_hex, end_color=wknd_hex, fill_type="solid")
        fill_header = PatternFill(start_color=colors["header"], end_color=colors["header"], fill_type="solid")

        # Phase header row
        if spec != current_spec:
            current_spec = spec
            ws.row_dimensions[row].height = 22
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            cell = ws.cell(row=row, column=1, value=spec)
            cell.font = phase_header_font
            cell.fill = fill_header
            cell.alignment = phase_align
            for c in range(1, 7):
                ws.cell(row=row, column=c).fill = fill_header
                ws.cell(row=row, column=c).border = thin_border
            for i, day in enumerate(all_days):
                col = DATA_COL_START + i
                ws.cell(row=row, column=col).fill = fill_header
                ws.cell(row=row, column=col).border = _border_for_day(day)
            row += 1

        # Task row
        ws.row_dimensions[row].height = 20
        task_data = [
            (1, spec, left_wrap),
            (2, spec, left_wrap),
            (3, task_name, left_wrap),
            (4, start_dt.strftime("%d.%m.%y"), center_align),
            (5, duration, center_align),
            (6, end_dt.strftime("%d.%m.%y"), center_align),
        ]
        for col, val, align in task_data:
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = task_font
            cell.fill = fill_phase
            cell.alignment = align
            cell.border = thin_border

        # Bars
        for i, day in enumerate(all_days):
            col = DATA_COL_START + i
            cell = ws.cell(row=row, column=col)
            cell.border = _border_for_day(day)
            if start_dt <= day <= end_dt and day.weekday() < 5:
                cell.fill = fill_bar
            elif day.weekday() >= 5:
                cell.fill = fill_wknd
            else:
                cell.fill = fill_phase

        row += 1

    ws.freeze_panes = "G4"


# ═══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python scripts/build_xlsx.py <decomposition.json> <output.xlsx> [--K 1.0] [--name \"Project Name\"]")
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
