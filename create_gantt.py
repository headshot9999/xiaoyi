"""
动态甘特图生成器

生成的 Excel 甘特图特性：
- 用户在"当前日期"单元格输入日期后，甘特图自动更新：
  - 完成度（%）根据日期自动计算
  - 状态（未开始/进行中/已完成）自动判断
  - 甘特条颜色通过条件格式自动变化（已过=蓝色，当前周=绿色，未来=浅蓝色）
  - "今天"标记线自动定位
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from datetime import datetime, date

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "项目甘特图"

# ── 样式定义 ──────────────────────────────────────────────

TITLE_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
TITLE_FONT = Font(name="微软雅黑", size=16, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
TASK_FONT = Font(name="微软雅黑", size=10)
TASK_FONT_BOLD = Font(name="微软雅黑", size=10, bold=True)
DATE_INPUT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
DATE_INPUT_FONT = Font(name="微软雅黑", size=12, bold=True, color="C00000")
LABEL_FONT = Font(name="微软雅黑", size=10, bold=True, color="1F4E79")

thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

STATUS_FILLS = {
    "已完成": PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid"),
    "进行中": PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid"),
    "未开始": PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid"),
}

FILL_DONE = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
FILL_CURRENT = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
FILL_PENDING = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
FILL_TODAY = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")

# ── 任务数据（使用具体日期范围） ────────────────────────────

tasks = [
    {
        "id": 1,
        "task": "供应商考察",
        "detail": "联系项目经理，到公司交流考察",
        "responsible": "采购部",
        "start": date(2026, 4, 1),
        "end": date(2026, 4, 30),
    },
    {
        "id": 2,
        "task": "合同签订",
        "detail": "完成合同签订流程",
        "responsible": "采购部/法务",
        "start": date(2026, 5, 1),
        "end": date(2026, 5, 31),
    },
    {
        "id": 3,
        "task": "开发工作",
        "detail": "系统开发与功能实现",
        "responsible": "开发部",
        "start": date(2026, 6, 1),
        "end": date(2026, 7, 31),
    },
    {
        "id": 4,
        "task": "系统测试及运营反馈",
        "detail": "系统测试、运营反馈与优化",
        "responsible": "测试/运营部",
        "start": date(2026, 8, 1),
        "end": date(2026, 9, 30),
    },
]

# ── 时间轴：按周划分（4月W1 ~ 9月W4） ─────────────────────

months = [4, 5, 6, 7, 8, 9]
YEAR = 2026
weeks_per_month = 4
month_names = {m: f"{m}月" for m in months}

# 计算每周的实际日期范围
week_dates = []
for m in months:
    if m == 9:
        last_day = 30
    elif m in (4, 6):
        last_day = 30
    else:
        last_day = 31
    month_days = last_day
    days_per_week = month_days / weeks_per_month
    for w in range(weeks_per_month):
        w_start = date(YEAR, m, int(w * days_per_week) + 1)
        w_end_day = min(int((w + 1) * days_per_week), last_day)
        w_end = date(YEAR, m, w_end_day)
        week_dates.append((w_start, w_end))

# ── 布局常量 ──────────────────────────────────────────────

INFO_COLS = 8  # 序号 | 任务 | 描述 | 负责 | 开始日期 | 结束日期 | 完成度 | 状态
COL_ID = 1
COL_TASK = 2
COL_DETAIL = 3
COL_RESPONSIBLE = 4
COL_START = 5
COL_END = 6
COL_PROGRESS = 7
COL_STATUS = 8
GANTT_START_COL = INFO_COLS + 1
TOTAL_WEEK_COLS = len(months) * weeks_per_month

# ── 第0行：当前日期输入区 ────────────────────────────────

INPUT_ROW = 1
ws.row_dimensions[INPUT_ROW].height = 35

ws.merge_cells(start_row=INPUT_ROW, start_column=1,
               end_row=INPUT_ROW, end_column=2)
label_cell = ws.cell(row=INPUT_ROW, column=1, value="📅 请输入当前日期：")
label_cell.font = LABEL_FONT
label_cell.alignment = Alignment(horizontal="right", vertical="center")

DATE_CELL = "C1"  # 用户在此输入日期
date_cell = ws.cell(row=INPUT_ROW, column=3)
date_cell.value = date.today()
date_cell.number_format = "YYYY-MM-DD"
date_cell.font = DATE_INPUT_FONT
date_cell.fill = DATE_INPUT_FILL
date_cell.alignment = Alignment(horizontal="center", vertical="center")
date_cell.border = Border(
    left=Side(style="medium", color="C00000"),
    right=Side(style="medium", color="C00000"),
    top=Side(style="medium", color="C00000"),
    bottom=Side(style="medium", color="C00000"),
)

ws.merge_cells(start_row=INPUT_ROW, start_column=4,
               end_row=INPUT_ROW, end_column=6)
hint_cell = ws.cell(row=INPUT_ROW, column=4,
                    value="← 修改此日期，甘特图自动更新")
hint_cell.font = Font(name="微软雅黑", size=9, italic=True, color="888888")
hint_cell.alignment = Alignment(horizontal="left", vertical="center")

# ── 第1行：标题 ───────────────────────────────────────────

TITLE_ROW = 2
ws.merge_cells(start_row=TITLE_ROW, start_column=1,
               end_row=TITLE_ROW, end_column=INFO_COLS + TOTAL_WEEK_COLS)
title_cell = ws.cell(row=TITLE_ROW, column=1)
title_cell.value = "项目甘特图 —— 2026年度项目推进计划"
title_cell.font = TITLE_FONT
title_cell.fill = TITLE_FILL
title_cell.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[TITLE_ROW].height = 40

# ── 第2-3行：表头（月份 + 周次） ──────────────────────────

ROW_MONTH = 3
ROW_WEEK = 4
ws.row_dimensions[ROW_MONTH].height = 25
ws.row_dimensions[ROW_WEEK].height = 22

info_headers = ["序号", "任务名称", "任务描述", "负责部门",
                "开始日期", "结束日期", "完成度", "状态"]
col_widths = [6, 18, 28, 14, 12, 12, 10, 10]

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

# ── 隐藏辅助行：存储每周的开始/结束日期 ──────────────────

AUX_ROW_START = ROW_WEEK + 1  # row 5: 周开始日期
AUX_ROW_END = ROW_WEEK + 2    # row 6: 周结束日期
ws.row_dimensions[AUX_ROW_START].height = 0.5
ws.row_dimensions[AUX_ROW_END].height = 0.5
ws.row_dimensions[AUX_ROW_START].hidden = True
ws.row_dimensions[AUX_ROW_END].hidden = True

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
        week_idx = mi * weeks_per_month + wi
        ws.column_dimensions[get_column_letter(col)].width = 5

        week_cell = ws.cell(row=ROW_WEEK, column=col, value=f"W{wi+1}")
        week_cell.font = Font(name="微软雅黑", size=8, bold=True, color="FFFFFF")
        week_cell.fill = PatternFill(start_color="4472C4", end_color="4472C4",
                                     fill_type="solid")
        week_cell.alignment = Alignment(horizontal="center", vertical="center")
        week_cell.border = thin_border

        ws.cell(row=AUX_ROW_START, column=col, value=week_dates[week_idx][0])
        ws.cell(row=AUX_ROW_END, column=col, value=week_dates[week_idx][1])

# ── 数据行 ────────────────────────────────────────────────

DATA_START_ROW = AUX_ROW_END + 1  # row 7

for idx, t in enumerate(tasks):
    row = DATA_START_ROW + idx
    ws.row_dimensions[row].height = 32

    even_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2",
                            fill_type="solid") if idx % 2 == 1 else None

    # 序号
    cell_id = ws.cell(row=row, column=COL_ID, value=t["id"])
    cell_id.font = TASK_FONT
    cell_id.alignment = Alignment(horizontal="center", vertical="center")
    cell_id.border = thin_border
    if even_fill:
        cell_id.fill = even_fill

    # 任务名称
    cell_task = ws.cell(row=row, column=COL_TASK, value=t["task"])
    cell_task.font = TASK_FONT_BOLD
    cell_task.alignment = Alignment(horizontal="left", vertical="center",
                                     wrap_text=True)
    cell_task.border = thin_border
    if even_fill:
        cell_task.fill = even_fill

    # 任务描述
    cell_detail = ws.cell(row=row, column=COL_DETAIL, value=t["detail"])
    cell_detail.font = TASK_FONT
    cell_detail.alignment = Alignment(horizontal="left", vertical="center",
                                       wrap_text=True)
    cell_detail.border = thin_border
    if even_fill:
        cell_detail.fill = even_fill

    # 负责部门
    cell_resp = ws.cell(row=row, column=COL_RESPONSIBLE, value=t["responsible"])
    cell_resp.font = TASK_FONT
    cell_resp.alignment = Alignment(horizontal="center", vertical="center")
    cell_resp.border = thin_border
    if even_fill:
        cell_resp.fill = even_fill

    # 开始日期
    cell_start = ws.cell(row=row, column=COL_START, value=t["start"])
    cell_start.number_format = "YYYY-MM-DD"
    cell_start.font = TASK_FONT
    cell_start.alignment = Alignment(horizontal="center", vertical="center")
    cell_start.border = thin_border
    if even_fill:
        cell_start.fill = even_fill

    # 结束日期
    cell_end = ws.cell(row=row, column=COL_END, value=t["end"])
    cell_end.number_format = "YYYY-MM-DD"
    cell_end.font = TASK_FONT
    cell_end.alignment = Alignment(horizontal="center", vertical="center")
    cell_end.border = thin_border
    if even_fill:
        cell_end.fill = even_fill

    # 完成度 —— 公式：根据当前日期自动计算
    # =IF($C$1<E7, 0, IF($C$1>=F7, 1, ($C$1-E7)/(F7-E7)))
    start_ref = f"${get_column_letter(COL_START)}${row}"
    end_ref = f"${get_column_letter(COL_END)}${row}"
    progress_formula = (
        f'=IF($C$1="",0,'
        f'IF($C$1<{start_ref},0,'
        f'IF($C$1>={end_ref},1,'
        f'($C$1-{start_ref})/({end_ref}-{start_ref}))))'
    )
    cell_prog = ws.cell(row=row, column=COL_PROGRESS)
    cell_prog.value = progress_formula
    cell_prog.number_format = "0%"
    cell_prog.font = TASK_FONT_BOLD
    cell_prog.alignment = Alignment(horizontal="center", vertical="center")
    cell_prog.border = thin_border

    # 状态 —— 公式：根据完成度自动判断
    prog_ref = f"${get_column_letter(COL_PROGRESS)}${row}"
    status_formula = (
        f'=IF({prog_ref}>=1,"已完成",'
        f'IF({prog_ref}>0,"进行中","未开始"))'
    )
    cell_status = ws.cell(row=row, column=COL_STATUS)
    cell_status.value = status_formula
    cell_status.font = TASK_FONT
    cell_status.alignment = Alignment(horizontal="center", vertical="center")
    cell_status.border = thin_border

    # 甘特条区域 —— 每个单元格用公式返回标记值
    # 值含义：0=不在范围, 1=已完成, 2=当前周, 3=未来, 4=今天所在周
    for mi, m in enumerate(months):
        for wi in range(weeks_per_month):
            col = GANTT_START_COL + mi * weeks_per_month + wi
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border

            ws_col = get_column_letter(col)
            # 辅助行中该周的开始/结束日期
            week_start_ref = f"{ws_col}${AUX_ROW_START}"
            week_end_ref = f"{ws_col}${AUX_ROW_END}"

            # 判断该周是否在任务范围内，以及相对于当前日期的状态
            # 0: 不在任务范围
            # 1: 已过去（已完成）
            # 2: 当前周（进行中）
            # 3: 未来（待完成）
            gantt_formula = (
                f'=IF(OR({week_start_ref}>{end_ref},{week_end_ref}<{start_ref}),0,'
                f'IF($C$1="",3,'
                f'IF($C$1>{week_end_ref},1,'
                f'IF(AND($C$1>={week_start_ref},$C$1<={week_end_ref}),2,3))))'
            )
            cell.value = gantt_formula
            cell.font = Font(size=1, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # ── 条件格式（对该行的甘特区域） ─────────────────────

    gantt_range = (
        f"{get_column_letter(GANTT_START_COL)}{row}:"
        f"{get_column_letter(GANTT_START_COL + TOTAL_WEEK_COLS - 1)}{row}"
    )

    # 值=1 → 已完成（蓝色）
    ws.conditional_formatting.add(
        gantt_range,
        CellIsRule(operator="equal", formula=["1"], fill=FILL_DONE)
    )
    # 值=2 → 当前周（绿色）
    ws.conditional_formatting.add(
        gantt_range,
        CellIsRule(operator="equal", formula=["2"], fill=FILL_CURRENT)
    )
    # 值=3 → 未来（浅蓝）
    ws.conditional_formatting.add(
        gantt_range,
        CellIsRule(operator="equal", formula=["3"], fill=FILL_PENDING)
    )

# ── 状态列条件格式 ───────────────────────────────────────

status_range = (
    f"{get_column_letter(COL_STATUS)}{DATA_START_ROW}:"
    f"{get_column_letter(COL_STATUS)}{DATA_START_ROW + len(tasks) - 1}"
)
ws.conditional_formatting.add(
    status_range,
    CellIsRule(operator="equal", formula=['"已完成"'],
               fill=STATUS_FILLS["已完成"],
               font=Font(name="微软雅黑", size=9, color="FFFFFF"))
)
ws.conditional_formatting.add(
    status_range,
    CellIsRule(operator="equal", formula=['"进行中"'],
               fill=STATUS_FILLS["进行中"],
               font=Font(name="微软雅黑", size=9, color="FFFFFF"))
)
ws.conditional_formatting.add(
    status_range,
    CellIsRule(operator="equal", formula=['"未开始"'],
               fill=STATUS_FILLS["未开始"],
               font=Font(name="微软雅黑", size=9, color="333333"))
)

# ── 完成度列条件格式：进度越高颜色越深 ─────────────────

progress_range = (
    f"{get_column_letter(COL_PROGRESS)}{DATA_START_ROW}:"
    f"{get_column_letter(COL_PROGRESS)}{DATA_START_ROW + len(tasks) - 1}"
)
ws.conditional_formatting.add(
    progress_range,
    FormulaRule(
        formula=[f"{get_column_letter(COL_PROGRESS)}{DATA_START_ROW}>=1"],
        fill=PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid"),
        font=Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    )
)
ws.conditional_formatting.add(
    progress_range,
    FormulaRule(
        formula=[f"AND({get_column_letter(COL_PROGRESS)}{DATA_START_ROW}>0,"
                 f"{get_column_letter(COL_PROGRESS)}{DATA_START_ROW}<1)"],
        fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),
        font=Font(name="微软雅黑", size=10, bold=True, color="2E75B6")
    )
)

# ── "今天"标记线：在甘特区域的周表头行用条件格式高亮 ──

for mi, m in enumerate(months):
    for wi in range(weeks_per_month):
        col = GANTT_START_COL + mi * weeks_per_month + wi
        ws_col = get_column_letter(col)
        week_start_ref = f"{ws_col}${AUX_ROW_START}"
        week_end_ref = f"{ws_col}${AUX_ROW_END}"

        cell_ref = f"{ws_col}{ROW_WEEK}"
        ws.conditional_formatting.add(
            cell_ref,
            FormulaRule(
                formula=[f"AND($C$1>={week_start_ref},$C$1<={week_end_ref})"],
                fill=PatternFill(start_color="FF6B6B", end_color="FF6B6B",
                                 fill_type="solid"),
                font=Font(name="微软雅黑", size=8, bold=True, color="FFFFFF")
            )
        )

# ── 图例 ─────────────────────────────────────────────────

LEGEND_ROW = DATA_START_ROW + len(tasks) + 2
ws.cell(row=LEGEND_ROW, column=1, value="图例：").font = Font(
    name="微软雅黑", size=10, bold=True)

legend_items = [
    ("已完成", FILL_DONE),
    ("当前周", FILL_CURRENT),
    ("未开始", FILL_PENDING),
    ("今天", FILL_TODAY),
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

# ── 使用说明 ──────────────────────────────────────────────

NOTE_ROW = LEGEND_ROW + 2
ws.merge_cells(start_row=NOTE_ROW, start_column=1,
               end_row=NOTE_ROW + 2, end_column=INFO_COLS + TOTAL_WEEK_COLS)
note_cell = ws.cell(row=NOTE_ROW, column=1)
note_cell.value = (
    "使用说明：\n"
    "1. 修改 C1 单元格的日期（格式：YYYY-MM-DD），甘特图会自动更新进度展示。\n"
    "2. 完成度根据日期在 [开始日期, 结束日期] 区间内的比例自动计算。\n"
    "3. 状态根据完成度自动判断：0%=未开始，0%~100%=进行中，100%=已完成。\n"
    "4. 甘特条颜色：蓝色=已过去的周（已完成），绿色=当前所在周，浅蓝色=未来的周。\n"
    '5. 红色标记表示"今天"所在的周次。'
)
note_cell.font = Font(name="微软雅黑", size=9, color="666666")
note_cell.alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[NOTE_ROW].height = 20
ws.row_dimensions[NOTE_ROW + 1].height = 20
ws.row_dimensions[NOTE_ROW + 2].height = 20

# ── 页面设置 ──────────────────────────────────────────────

ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(
    fitToPage=True
)
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0

ws.freeze_panes = f"{get_column_letter(GANTT_START_COL)}{DATA_START_ROW}"

output_path = "/workspace/项目甘特图.xlsx"
wb.save(output_path)
print(f"✅ 动态甘特图已生成: {output_path}")
print(f"   打开 Excel 后修改 C1 单元格的日期即可自动更新进度展示。")
