#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
证件照 Word 排版生成器
=====================

将一张或多张照片处理成标准证件照尺寸（默认 1 寸 = 2.5cm × 3.5cm），
并在 Word 文档（A4 或 5 寸相纸）中排成网格，可直接打印后裁剪使用。

特性：
  * 自动裁剪到证件照比例（默认居中裁剪，不变形）；
  * 在 Word 中以"物理厘米尺寸"精确插入，打印即真实尺寸；
  * 网格排版 + 可选裁剪边线，方便裁切；
  * 支持 1 寸 / 小 1 寸 / 2 寸 / 小 2 寸 等常见规格，也支持自定义尺寸。

用法示例：
  # 处理 photos 目录下所有图片，每张照片各排满一页 A4
  python make_photo_doc.py -i photos -o 证件照.docx

  # 指定多张图片，每张冲印 8 张，使用 5 寸相纸版式
  python make_photo_doc.py -i a.jpg b.jpg --page 5r --copies 8

  # 2 寸照片、不画裁剪线
  python make_photo_doc.py -i photo.jpg --size 2inch --no-border
"""

import argparse
import io
import os
import sys

from PIL import Image, ImageOps

from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ---------------------------------------------------------------------------
# 规格定义（单位：毫米，宽 × 高）
# ---------------------------------------------------------------------------
PHOTO_SIZES = {
    "1inch": (25, 35),       # 标准 1 寸
    "small1inch": (22, 32),  # 小 1 寸（第二代身份证常用）
    "2inch": (35, 49),       # 标准 2 寸
    "small2inch": (35, 45),  # 小 2 寸（护照 / 签证常用）
    "id": (26, 32),          # 二代身份证
}

# 纸张可打印区域（单位：毫米，宽 × 高）以及推荐页边距
PAGE_SIZES = {
    "a4": (210, 297),
    "5r": (127, 89),   # 5 寸相纸（横向 7×5），一版多张证件照常用
    "6r": (152, 102),  # 6 寸相纸
}


def parse_args():
    p = argparse.ArgumentParser(
        description="将照片排版成可打印的证件照 Word 文档",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument(
        "-i", "--input", nargs="+", required=True,
        help="输入图片文件或目录（可多个）",
    )
    p.add_argument(
        "-o", "--output", default="证件照.docx",
        help="输出 Word 文件名（默认：证件照.docx）",
    )
    p.add_argument(
        "--size", default="1inch",
        help="照片规格：1inch/small1inch/2inch/small2inch/id，"
             "或自定义 '宽x高'（毫米，如 30x40）。默认 1inch",
    )
    p.add_argument(
        "--page", default="a4", choices=list(PAGE_SIZES.keys()),
        help="纸张：a4 / 5r / 6r（默认 a4）",
    )
    p.add_argument(
        "--copies", type=int, default=0,
        help="每张照片打印的份数；0 表示用照片填满整页（默认 0）",
    )
    p.add_argument(
        "--gap", type=float, default=2.0,
        help="照片之间的间距（毫米，默认 2）",
    )
    p.add_argument(
        "--margin", type=float, default=8.0,
        help="页边距（毫米，默认 8）",
    )
    p.add_argument(
        "--no-border", action="store_true",
        help="不绘制裁剪边线",
    )
    p.add_argument(
        "--fit", action="store_true",
        help="保留整张照片（留白），而不是居中裁剪到证件照比例",
    )
    p.add_argument(
        "--dpi", type=int, default=300,
        help="重采样目标 DPI（默认 300，用于决定图片像素分辨率）",
    )
    return p.parse_args()


def resolve_size(spec):
    """解析照片规格 -> (宽mm, 高mm)。"""
    if spec in PHOTO_SIZES:
        return PHOTO_SIZES[spec]
    if "x" in spec.lower():
        try:
            w, h = spec.lower().split("x")
            return float(w), float(h)
        except ValueError:
            pass
    raise SystemExit(f"无法识别的尺寸规格：{spec}")


def collect_images(inputs):
    """把文件 / 目录展开成图片文件列表。"""
    exts = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff"}
    files = []
    for item in inputs:
        if os.path.isdir(item):
            for name in sorted(os.listdir(item)):
                if os.path.splitext(name)[1].lower() in exts:
                    files.append(os.path.join(item, name))
        elif os.path.isfile(item):
            files.append(item)
        else:
            print(f"  ⚠ 跳过：找不到 {item}", file=sys.stderr)
    return files


def prepare_image(path, size_mm, dpi, fit):
    """读取图片，按证件照比例处理，返回内存中的 PNG 字节流。"""
    img = Image.open(path)
    img = ImageOps.exif_transpose(img)  # 修正手机拍摄方向
    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGB")

    w_mm, h_mm = size_mm
    target_px = (
        max(1, round(w_mm / 25.4 * dpi)),
        max(1, round(h_mm / 25.4 * dpi)),
    )

    if fit:
        # 保留整张图，按比例缩放后用白底填充到目标比例
        bg = Image.new("RGB", target_px, "white")
        scaled = ImageOps.contain(img.convert("RGB"), target_px)
        off = ((target_px[0] - scaled.width) // 2,
               (target_px[1] - scaled.height) // 2)
        bg.paste(scaled, off)
        out = bg
    else:
        # 居中裁剪到证件照比例
        out = ImageOps.fit(img.convert("RGB"), target_px,
                           method=Image.LANCZOS, centering=(0.5, 0.5))

    buf = io.BytesIO()
    out.save(buf, format="PNG", dpi=(dpi, dpi))
    buf.seek(0)
    return buf


def set_cell_border(cell, color="999999", sz=4):
    """给单元格四周加细边线（作为裁剪参考线）。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tc_pr.append(borders)


def set_row_height(row, height_mm):
    """精确设置行高（exact）。"""
    tr_pr = row._tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(int(height_mm / 25.4 * 1440)))  # twips
    h.set(qn("w:hRule"), "exact")
    tr_pr.append(h)


def disable_autofit(table):
    tbl_pr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)


def set_cell_width(cell, width_mm):
    tc_pr = cell._tc.get_or_add_tcPr()
    w = OxmlElement("w:tcW")
    w.set(qn("w:w"), str(int(width_mm / 25.4 * 1440)))
    w.set(qn("w:type"), "dxa")
    tc_pr.append(w)


def set_cell_no_margins(cell):
    """去掉单元格内边距，保证照片精确尺寸。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tc_pr.append(mar)


def compute_grid(page, margin, photo, gap):
    """计算每页能排多少行多少列。"""
    pw, ph = page
    w_mm, h_mm = photo
    avail_w = pw - 2 * margin
    avail_h = ph - 2 * margin
    cols = max(1, int((avail_w + gap) // (w_mm + gap)))
    rows = max(1, int((avail_h + gap) // (h_mm + gap)))
    return cols, rows


def build_table(doc, slots, image_buffers, photo_mm, gap, border):
    """在文档当前页放置一个 cols×rows 的照片表格。slots 为 (cols, rows)。"""
    cols, rows = slots
    w_mm, h_mm = photo_mm
    table = doc.add_table(rows=rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    disable_autofit(table)

    idx = 0
    for r in range(rows):
        set_row_height(table.rows[r], h_mm)
        for c in range(cols):
            cell = table.cell(r, c)
            set_cell_width(cell, w_mm)
            set_cell_no_margins(cell)
            if border:
                set_cell_border(cell)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.line_spacing = 1.0
            if idx < len(image_buffers):
                buf = io.BytesIO(image_buffers[idx].getvalue())
                run = para.add_run()
                run.add_picture(buf, width=Mm(w_mm), height=Mm(h_mm))
            idx += 1
    return table


def main():
    args = parse_args()
    photo_mm = resolve_size(args.size)
    page_mm = PAGE_SIZES[args.page]

    files = collect_images(args.input)
    if not files:
        raise SystemExit("没有找到任何图片，请检查 -i 参数。")

    print(f"照片规格：{photo_mm[0]}mm × {photo_mm[1]}mm | 纸张：{args.page.upper()}")
    print(f"找到 {len(files)} 张图片：")
    for f in files:
        print(f"  - {f}")

    cols, rows = compute_grid(page_mm, args.margin, photo_mm, args.gap)
    per_page = cols * rows
    print(f"每页布局：{cols} 列 × {rows} 行 = {per_page} 张/页")

    # 处理图片
    prepared = []  # list of (filename, buffer)
    for f in files:
        try:
            buf = prepare_image(f, photo_mm, args.dpi, args.fit)
            prepared.append((os.path.basename(f), buf))
        except Exception as e:  # noqa: BLE001
            print(f"  ⚠ 处理失败 {f}: {e}", file=sys.stderr)

    if not prepared:
        raise SystemExit("没有可用的图片。")

    # 创建文档
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(page_mm[0])
    section.page_height = Mm(page_mm[1])
    section.top_margin = Mm(args.margin)
    section.bottom_margin = Mm(args.margin)
    section.left_margin = Mm(args.margin)
    section.right_margin = Mm(args.margin)

    first_page = True

    def new_page():
        nonlocal first_page
        if first_page:
            first_page = False
        else:
            doc.add_page_break()

    if args.copies and args.copies > 0:
        # 每张照片单独排版，重复 copies 份
        for name, buf in prepared:
            remaining = args.copies
            while remaining > 0:
                new_page()
                n = min(remaining, per_page)
                build_table(doc, (cols, rows),
                            [buf] * n, photo_mm, args.gap, not args.no_border)
                remaining -= n
    elif len(prepared) == 1:
        # 单张照片：填满一页
        name, buf = prepared[0]
        new_page()
        build_table(doc, (cols, rows),
                    [buf] * per_page, photo_mm, args.gap, not args.no_border)
    else:
        # 多张照片：每张各占一整页（填满）
        for name, buf in prepared:
            new_page()
            build_table(doc, (cols, rows),
                        [buf] * per_page, photo_mm, args.gap, not args.no_border)

    doc.save(args.output)
    print(f"\n✅ 已生成：{args.output}")
    print("   打印时请选择『实际大小 / 100%』，不要勾选『适应页面缩放』，"
          "以保证证件照尺寸准确。")


if __name__ == "__main__":
    main()
