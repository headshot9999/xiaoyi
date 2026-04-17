import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "项目甘特图"

TITLE_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
TITLE_FONT = Font(name="微软雅黑", size=16, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
TASK_FONT = Font(name="微软雅黑", size=10)
TASK_FONT_BOLD = Font(name="微软雅黑", size=10, bold=True)

COLORS = {
    "done": PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid"),
    "in_progress": PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid"),
    "pending": PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),
    "milestone": PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid"),
}

STATUS_FILL = {
    "已完成": PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid"),
    "进行中": PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid"),
    "未开始": PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid"),
}
STATUS_FONT_COLOR = {
    "已完成": Font(name="微软雅黑", size=9, color="FFFFFF"),
    "进行中": Font(name="微软雅黑", size=9, color="FFFFFF"),
    "未开始": Font(name="微软雅黑", size=9, color="333333"),
}

thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

tasks = [
    {
        "id": 1,
        "task": "供应商考察",
        "detail": "联系项目经理，到公司交流考察",
        "responsible": "采购部",
        "start_month": 4,
        "start_week": 1,
        "end_month": 4,
        "end_week": 4,
        "progress": 60,
        "status": "进行中",
    },
    {
        "id": 2,
        "task": "合同签订",
        "detail": "完成合同签订流程",
        "responsible": "采购部/法务",
        "start_month": 5,
        "start_week": 1,
        "end_month": 5,
        "end_week": 4,
        "progress": 0,
        "status": "未开始",
    },
    {
        "id": 3,
        "task": "开发工作",
        "detail": "系统开发与功能实现",
        "responsible": "开发部",
        "start_month": 6,
        "start_week": 1,
        "end_month": 7,
        "end_week": 4,
        "progress": 0,
        "status": "未开始",
    },
    {
        "id": 4,
        "task": "系统测试及运营反馈",
        "detail": "系统测试、运营反馈与优化",
        "responsible": "测试/运营部",
        "start_month": 8,
        "start_week": 1,
        "end_month": 9,
        "end_week": 4,
        "progress": 0,
        "status": "未开始",
    },
]

months = [4, 5, 6, 7, 8, 9]
weeks_per_month = 4
month_names = {4: "4月", 5: "5月", 6: "6月", 7: "7月", 8: "8月", 9: "9月"}

INFO_COLS = 6
COL_ID = 1
COL_TASK = 2
COL_DETAIL = 3
COL_RESPONSIBLE = 4
COL_PROGRESS = 5
COL_STATUS = 6
GANTT_START_COL = INFO_COLS + 1
TOTAL_WEEK_COLS = len(months) * weeks_per_month

ws.merge_cells(start_row=1, start_column=1,
               end_row=1, end_column=INFO_COLS + TOTAL_WEEK_COLS)
title_cell = ws.cell(row=1, column=1)
title_cell.value = "项目甘特图 —— 2026年度项目推进计划"
title_cell.font = TITLE_FONT
title_cell.fill = TITLE_FILL
title_cell.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 40

ROW_MONTH = 2
ROW_WEEK = 3
ws.row_dimensions[ROW_MONTH].height = 25
ws.row_dimensions[ROW_WEEK].height = 22

info_headers = ["序号", "任务名称", "任务描述", "负责部门", "完成度", "状态"]
col_widths = [6, 18, 28, 14, 10, 10]

for i, (header, width) in enumerate(zip(info_headers, col_widths)):
    col = i + 1
    ws.column_dimensions[get_column_letter(col)].width = width

    cell_m = ws.cell(row=ROW_MONTH, column=col, value=header)
    cell_m.font = HEADER_FONT
    cell_m.fill = HEADER_FILL
    cell_m.alignment = Alignment(horizontal="center", vertical="center",
                                  wrap_text=True)
    cell_m.border = thin_border

    cell_w = ws.cell(row=ROW_WEEK, column=col)
    cell_w.fill = HEADER_FILL
    cell_w.border = thin_border

    ws.merge_cells(start_row=ROW_MONTH, start_column=col,
                   end_row=ROW_WEEK, end_column=col)

for mi, m in enumerate(months):
    start_col = GANTT_START_COL + mi * weeks_per_month
    end_col = start_col + weeks_per_month - 1

    ws.merge_cells(start_row=ROW_MONTH, start_column=start_col,
                   end_row=ROW_MONTH, end_column=end_col)
    month_cell = ws.cell(row=ROW_MONTH, column=start_col, value=month_names[m])
    month_cell.font = HEADER_FONT
    month_cell.fill = HEADER_FILL
    month_cell.alignment = Alignment(horizontal="center", vertical="center")
    month_cell.border = thin_border

    for wi in range(weeks_per_month):
        col = start_col + wi
        ws.column_dimensions[get_column_letter(col)].width = 5
        week_cell = ws.cell(row=ROW_WEEK, column=col, value=f"W{wi+1}")
        week_cell.font = Font(name="微软雅黑", size=8, bold=True, color="FFFFFF")
        week_cell.fill = PatternFill(start_color="4472C4", end_color="4472C4",
                                     fill_type="solid")
        week_cell.alignment = Alignment(horizontal="center", vertical="center")
        week_cell.border = thin_border

DATA_START_ROW = 4

for idx, t in enumerate(tasks):
    row = DATA_START_ROW + idx
    ws.row_dimensions[row].height = 32

    even_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2",
                            fill_type="solid") if idx % 2 == 1 else None

    cell_id = ws.cell(row=row, column=COL_ID, value=t["id"])
    cell_id.font = TASK_FONT
    cell_id.alignment = Alignment(horizontal="center", vertical="center")
    cell_id.border = thin_border
    if even_fill:
        cell_id.fill = even_fill

    cell_task = ws.cell(row=row, column=COL_TASK, value=t["task"])
    cell_task.font = TASK_FONT_BOLD
    cell_task.alignment = Alignment(horizontal="left", vertical="center",
                                     wrap_text=True)
    cell_task.border = thin_border
    if even_fill:
        cell_task.fill = even_fill

    cell_detail = ws.cell(row=row, column=COL_DETAIL, value=t["detail"])
    cell_detail.font = TASK_FONT
    cell_detail.alignment = Alignment(horizontal="left", vertical="center",
                                       wrap_text=True)
    cell_detail.border = thin_border
    if even_fill:
        cell_detail.fill = even_fill

    cell_resp = ws.cell(row=row, column=COL_RESPONSIBLE, value=t["responsible"])
    cell_resp.font = TASK_FONT
    cell_resp.alignment = Alignment(horizontal="center", vertical="center")
    cell_resp.border = thin_border
    if even_fill:
        cell_resp.fill = even_fill

    cell_prog = ws.cell(row=row, column=COL_PROGRESS, value=t["progress"] / 100)
    cell_prog.font = TASK_FONT_BOLD
    cell_prog.number_format = "0%"
    cell_prog.alignment = Alignment(horizontal="center", vertical="center")
    cell_prog.border = thin_border
    if t["progress"] > 0:
        cell_prog.font = Font(name="微软雅黑", size=10, bold=True, color="2E75B6")
    if even_fill:
        cell_prog.fill = even_fill

    cell_status = ws.cell(row=row, column=COL_STATUS, value=t["status"])
    cell_status.font = STATUS_FONT_COLOR.get(t["status"], TASK_FONT)
    cell_status.fill = STATUS_FILL.get(t["status"],
                                        PatternFill(fill_type=None))
    cell_status.alignment = Alignment(horizontal="center", vertical="center")
    cell_status.border = thin_border

    for mi, m in enumerate(months):
        for wi in range(weeks_per_month):
            col = GANTT_START_COL + mi * weeks_per_month + wi
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            if even_fill:
                cell.fill = even_fill

            global_week = (m - months[0]) * weeks_per_month + wi
            task_start = (t["start_month"] - months[0]) * weeks_per_month + (t["start_week"] - 1)
            task_end = (t["end_month"] - months[0]) * weeks_per_month + (t["end_week"] - 1)

            if task_start <= global_week <= task_end:
                total_weeks = task_end - task_start + 1
                progress_weeks = int(total_weeks * t["progress"] / 100)
                week_in_task = global_week - task_start

                if t["progress"] > 0 and week_in_task < progress_weeks:
                    cell.fill = COLORS["done"]
                elif t["progress"] > 0 and week_in_task == progress_weeks:
                    cell.fill = COLORS["in_progress"]
                else:
                    cell.fill = COLORS["pending"]

LEGEND_ROW = DATA_START_ROW + len(tasks) + 2
ws.cell(row=LEGEND_ROW, column=1, value="图例：").font = Font(
    name="微软雅黑", size=10, bold=True)

legend_items = [
    ("已完成", COLORS["done"]),
    ("进行中", COLORS["in_progress"]),
    ("未开始", COLORS["pending"]),
]

legend_col = 2
for label, fill in legend_items:
    cell_color = ws.cell(row=LEGEND_ROW, column=legend_col)
    cell_color.fill = fill
    cell_color.border = thin_border
    legend_col += 1

    cell_label = ws.cell(row=LEGEND_ROW, column=legend_col, value=label)
    cell_label.font = Font(name="微软雅黑", size=9)
    cell_label.alignment = Alignment(vertical="center")
    legend_col += 1

NOTE_ROW = LEGEND_ROW + 2
ws.merge_cells(start_row=NOTE_ROW, start_column=1,
               end_row=NOTE_ROW, end_column=INFO_COLS + TOTAL_WEEK_COLS)
note_cell = ws.cell(row=NOTE_ROW, column=1)
note_cell.value = (
    "备注：4月完成供应商考察（目前已联系项目经理，预计本周内回复，"
    "下周可到公司交流，当前完成度60%）；"
    "5月完成合同签订；6-7月完成开发工作；8-9月系统测试及运营反馈。"
)
note_cell.font = Font(name="微软雅黑", size=9, color="666666")
note_cell.alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[NOTE_ROW].height = 40

ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(
    fitToPage=True
)
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0

ws.freeze_panes = "G4"

output_path = "/workspace/项目甘特图.xlsx"
wb.save(output_path)
print(f"甘特图已生成: {output_path}")
