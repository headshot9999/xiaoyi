#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
艾多美 5 周年员工旅游心得 PPT 生成器
====================================

根据员工 5 周年全家沪杭之旅（上海外滩、金茂大厦 56 层观景自助、豫园、城隍庙、南京路、
杭州西湖、雷峰塔、海鲜大餐等）自动生成图文并茂的旅游心得演示文稿（.pptx）。

用法：
  # 直接运行（自动读取默认照片目录：/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片）
  python3 make_travel_ppt.py

  # 或手动指定照片目录与输出文件名
  python3 make_travel_ppt.py -p "/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片" -o 艾多美5周年员工旅游心得.pptx
"""

from __future__ import annotations

import argparse
import io
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple

from PIL import Image, ImageDraw, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

try:
    from pillow_heif import register_heif_opener

    register_heif_opener()
except Exception:
    pass


DEFAULT_PHOTO_DIR = "/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片"
DEFAULT_OUTPUT = "艾多美5周年员工旅游心得.pptx"

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".heic", ".heif"}

# 配色方案（艾多美蓝 + 温暖金/米白杂志风）
C_BG = "F8FAFC"          # 浅云白背景
C_CARD = "FFFFFF"        # 白色卡片
C_NAVY = "0F2942"        # 深海蓝主标题
C_ATOMY = "0095D9"       # 艾多美天蓝
C_ATOMY_SOFT = "E6F4FB"  # 浅蓝标签底色
C_GOLD = "C88A2E"        # 纪念暖金
C_GOLD_SOFT = "FDF5E6"   # 暖金浅底
C_TEXT = "1E293B"        # 正文深灰蓝
C_MUTED = "64748B"       # 辅助灰字
C_BORDER = "E2E8F0"      # 细边框

FONT_MAIN = "PingFang SC"
FONT_FALLBACK = "Microsoft YaHei"

SLIDE_W = 13.333
SLIDE_H = 7.5


def rgb(hex_str: str) -> RGBColor:
    return RGBColor.from_string(hex_str.lstrip("#"))


def set_run_font(
    run,
    size: float,
    color_hex: str,
    bold: bool = False,
    font_name: str = FONT_MAIN,
) -> None:
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color_hex)
    run.font.name = font_name
    rpr = run._r.get_or_add_rPr()
    rpr.set("lang", "zh-CN")
    rpr.set("altLang", "en-US")
    for tag in ("latin", "ea", "cs"):
        el = rpr.find(qn(f"a:{tag}"))
        if el is None:
            el = rpr.makeelement(qn(f"a:{tag}"), {"typeface": font_name})
            rpr.append(el)
        else:
            el.set("typeface", font_name)


def add_bg(slide, color_hex: str = C_BG) -> None:
    rect = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(SLIDE_W), Inches(SLIDE_H)
    )
    rect.fill.solid()
    rect.fill.fore_color.rgb = rgb(color_hex)
    rect.line.fill.background()


def add_box(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_hex: Optional[str] = C_CARD,
    border_hex: Optional[str] = None,
    border_pt: float = 1.0,
    rounded: bool = True,
):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(
        shape_type, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    if rounded and shape.adjustments:
        shape.adjustments[0] = 0.03
    if fill_hex:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill_hex)
    else:
        shape.fill.background()
    if border_hex:
        shape.line.color.rgb = rgb(border_hex)
        shape.line.width = Pt(border_pt)
    else:
        shape.line.fill.background()
    return shape


def add_text(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    size: float = 14,
    color_hex: str = C_TEXT,
    bold: bool = False,
    align=PP_ALIGN.LEFT,
    valign=MSO_ANCHOR.TOP,
    line_spacing: float = 1.3,
):
    tb = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)

    lines = text.split("\n")
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        if idx > 0:
            p.space_before = Pt(4)
        run = p.add_run()
        run.text = line
        set_run_font(run, size=size, color_hex=color_hex, bold=bold)
    return tb


def add_pill(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    bg_hex: str = C_ATOMY_SOFT,
    fg_hex: str = C_ATOMY,
    size: float = 11,
):
    pill = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    if pill.adjustments:
        pill.adjustments[0] = 0.5
    pill.fill.solid()
    pill.fill.fore_color.rgb = rgb(bg_hex)
    pill.line.fill.background()
    tf = pill.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    set_run_font(run, size=size, color_hex=fg_hex, bold=True)
    return pill


def add_header(
    slide,
    page_num: str,
    tag_text: str,
    title: str,
    subtitle: str,
) -> None:
    # 顶部细装饰条
    add_box(slide, 0, 0, SLIDE_W, 0.08, fill_hex=C_ATOMY, rounded=False)

    add_pill(slide, 0.65, 0.35, 1.85, 0.36, tag_text, bg_hex=C_ATOMY_SOFT, fg_hex=C_ATOMY, size=11)
    add_text(
        slide,
        2.62,
        0.31,
        8.5,
        0.48,
        title,
        size=22,
        color_hex=C_NAVY,
        bold=True,
    )
    add_text(
        slide,
        0.65,
        0.82,
        11.2,
        0.36,
        subtitle,
        size=12.5,
        color_hex=C_MUTED,
        bold=False,
    )
    # 右上页码与主题标识
    add_text(
        slide,
        10.8,
        0.36,
        1.9,
        0.35,
        f"ATOMY 5th  |  {page_num}",
        size=11,
        color_hex=C_MUTED,
        bold=True,
        align=PP_ALIGN.RIGHT,
    )


# ---------------------------------------------------------------------------
# 照片扫描、智能匹配与居中裁剪
# ---------------------------------------------------------------------------
class PhotoPool:
    def __init__(self, folders: Iterable[Path]):
        self.photos: List[Path] = []
        seen = set()
        for folder in folders:
            if not folder.exists() or not folder.is_dir():
                continue
            for p in sorted(folder.rglob("*")):
                if (
                    p.is_file()
                    and p.suffix.lower() in IMAGE_EXTS
                    and not p.name.startswith(".")
                ):
                    resolved = p.resolve()
                    if resolved not in seen:
                        seen.add(resolved)
                        self.photos.append(p)
        self.used: set[Path] = set()
        self.cursor: int = 0

    def pick(self, keywords: Sequence[str] = ()) -> Optional[Path]:
        if not self.photos:
            return None
        # 1. 优先按文件名或父目录名包含关键词匹配（未使用的优先）
        for kw in keywords:
            kw_low = kw.lower()
            for p in self.photos:
                if p not in self.used and (
                    kw_low in p.stem.lower() or kw_low in str(p.parent).lower()
                ):
                    self.used.add(p)
                    return p
        # 2. 按顺序取下一张未使用的照片
        for p in self.photos:
            if p not in self.used:
                self.used.add(p)
                return p
        # 3. 照片数少于占位框时循环使用
        p = self.photos[self.cursor % len(self.photos)]
        self.cursor += 1
        return p


def render_cropped_image_bytes(
    img_path: Path, target_w_in: float, target_h_in: float, dpi: int = 220
) -> io.BytesIO:
    target_px = (max(int(target_w_in * dpi), 300), max(int(target_h_in * dpi), 200))
    with Image.open(img_path) as im:
        im = ImageOps.exif_transpose(im)
        if im.mode not in ("RGB", "RGBA"):
            im = im.convert("RGB")
        fitted = ImageOps.fit(
            im.convert("RGB"), target_px, method=Image.Resampling.LANCZOS
        )
        buf = io.BytesIO()
        fitted.save(buf, format="JPEG", quality=90)
        buf.seek(0)
        return buf


def render_placeholder_bytes(
    target_w_in: float,
    target_h_in: float,
    theme: str = "bund",
    dpi: int = 160,
) -> io.BytesIO:
    """当本地还未放入照片时，生成优雅的风景渐变插画底图，方便直接预览与右键换图。"""
    w = max(int(target_w_in * dpi), 320)
    h = max(int(target_h_in * dpi), 220)
    palettes = {
        "family": ((18, 53, 88), (45, 125, 175), (242, 198, 132)),
        "bund": ((15, 34, 64), (35, 84, 138), (235, 180, 98)),
        "jinmao": ((28, 42, 74), (82, 112, 158), (244, 204, 138)),
        "yuyuan": ((92, 34, 38), (156, 68, 58), (240, 202, 145)),
        "nanjing": ((26, 44, 78), (118, 72, 126), (242, 184, 118)),
        "westlake": ((22, 68, 78), (64, 136, 132), (214, 232, 220)),
        "leifeng": ((58, 46, 72), (148, 96, 84), (244, 194, 122)),
        "seafood": ((112, 38, 32), (192, 82, 48), (248, 208, 136)),
    }
    c_top, c_mid, c_bot = palettes.get(theme, palettes["bund"])
    img = Image.new("RGB", (w, h))
    draw = ImageDraw.Draw(img)

    for y in range(h):
        t = y / max(h - 1, 1)
        if t < 0.55:
            k = t / 0.55
            r = int(c_top[0] + (c_mid[0] - c_top[0]) * k)
            g = int(c_top[1] + (c_mid[1] - c_top[1]) * k)
            b = int(c_top[2] + (c_mid[2] - c_top[2]) * k)
        else:
            k = (t - 0.55) / 0.45
            r = int(c_mid[0] + (c_bot[0] - c_mid[0]) * k)
            g = int(c_mid[1] + (c_bot[1] - c_mid[1]) * k)
            b = int(c_mid[2] + (c_bot[2] - c_mid[2]) * k)
        draw.line([(0, y), (w, y)], fill=(r, g, b))

    # 柔和太阳/明月与远山剪影
    sun_r = int(min(w, h) * 0.13)
    sun_x, sun_y = int(w * 0.76), int(h * 0.30)
    draw.ellipse(
        (sun_x - sun_r, sun_y - sun_r, sun_x + sun_r, sun_y + sun_r),
        fill=(252, 234, 192),
    )

    # 远山/城市天际线装饰
    horizon_y = int(h * 0.74)
    if theme in ("bund", "jinmao", "nanjing"):
        # 东方明珠/金茂天际线剪影
        buildings = [
            (0.06, 0.58, 0.07),
            (0.15, 0.48, 0.06),
            (0.24, 0.34, 0.07),
            (0.35, 0.52, 0.06),
            (0.45, 0.26, 0.08),
            (0.57, 0.42, 0.07),
            (0.68, 0.55, 0.08),
            (0.80, 0.46, 0.07),
        ]
        for bx, by, bw in buildings:
            x0 = int(w * bx)
            y0 = int(h * by)
            x1 = int(w * (bx + bw))
            draw.rectangle([x0, y0, x1, horizon_y], fill=(20, 34, 56))
    else:
        # 江南远山与水波
        pts = [
            (0, horizon_y),
            (int(w * 0.18), int(h * 0.56)),
            (int(w * 0.36), int(h * 0.65)),
            (int(w * 0.58), int(h * 0.50)),
            (int(w * 0.78), int(h * 0.62)),
            (w, int(h * 0.54)),
            (w, horizon_y),
        ]
        draw.polygon(pts, fill=(28, 52, 68))

    draw.rectangle([0, horizon_y, w, h], fill=(18, 36, 54))
    for i in range(4):
        wy = horizon_y + int((h - horizon_y) * (0.2 + i * 0.2))
        draw.line(
            [(int(w * 0.15), wy), (int(w * 0.85), wy)],
            fill=(240, 210, 156),
            width=2,
        )

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=88)
    buf.seek(0)
    return buf


def add_photo_frame(
    slide,
    pool: PhotoPool,
    left: float,
    top: float,
    width: float,
    height: float,
    caption: str,
    keywords: Sequence[str] = (),
    theme: str = "bund",
) -> None:
    """绘制带白色拍立得相框、照片（或插画占位）与底部文字图注的卡片。"""
    pad = 0.08
    cap_h = 0.42
    add_box(
        slide,
        left,
        top,
        width,
        height,
        fill_hex=C_CARD,
        border_hex=C_BORDER,
        border_pt=1.0,
        rounded=True,
    )

    img_left = left + pad
    img_top = top + pad
    img_w = width - pad * 2
    img_h = height - pad * 2 - cap_h

    chosen = pool.pick(keywords)
    img_stream: Optional[io.BytesIO] = None
    if chosen is not None:
        try:
            img_stream = render_cropped_image_bytes(chosen, img_w, img_h)
        except Exception:
            img_stream = None

    if img_stream is None:
        img_stream = render_placeholder_bytes(img_w, img_h, theme=theme)

    slide.shapes.add_picture(
        img_stream,
        Inches(img_left),
        Inches(img_top),
        Inches(img_w),
        Inches(img_h),
    )

    if chosen is None:
        # 提示标签：如果是在未加载真实照片的环境下生成，右上角给出温和提示
        add_pill(
            slide,
            img_left + 0.12,
            img_top + 0.12,
            min(2.3, img_w - 0.24),
            0.30,
            f"照片位：{caption.split('｜')[0]}",
            bg_hex="1E293B",
            fg_hex="F8FAFC",
            size=9.5,
        )

    add_text(
        slide,
        left + 0.12,
        top + height - cap_h - 0.02,
        width - 0.24,
        cap_h,
        caption,
        size=11,
        color_hex=C_NAVY,
        bold=True,
        align=PP_ALIGN.CENTER,
        valign=MSO_ANCHOR.MIDDLE,
    )


# ---------------------------------------------------------------------------
# 各页幻灯片构建（共 7 页）
# ---------------------------------------------------------------------------
def build_slide_1_cover(prs: Presentation, pool: PhotoPool) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, C_NAVY)

    # 顶部与左侧装饰色块
    add_box(slide, 0, 0, SLIDE_W, 0.10, fill_hex=C_ATOMY, rounded=False)
    add_box(slide, 0.72, 0.85, 6.25, 5.85, fill_hex="163756", border_hex="275279", rounded=True)

    add_pill(
        slide,
        1.05,
        1.20,
        2.95,
        0.42,
        "ATOMY · 入职满 5 周年员工旅游",
        bg_hex=C_ATOMY,
        fg_hex="FFFFFF",
        size=12,
    )

    add_text(
        slide,
        1.05,
        1.85,
        5.6,
        1.45,
        "感恩五载同行\n携家人共赴沪杭美好时光",
        size=30,
        color_hex="FFFFFF",
        bold=True,
        line_spacing=1.25,
    )

    add_text(
        slide,
        1.05,
        3.40,
        5.5,
        0.85,
        "上海（外滩 · 金茂大厦56层 · 豫园 · 城隍庙 · 南京路）\n杭州（西湖 · 雷峰塔 · 海鲜盛宴）",
        size=13.5,
        color_hex="CBD5E1",
        line_spacing=1.4,
    )

    # 里程碑徽章（3 / 5 / 7 / 10 年）
    years = [("3年", False), ("5年 · 本次", True), ("7年", False), ("10年", False)]
    for i, (label, active) in enumerate(years):
        bx = 1.05 + i * 1.38
        add_pill(
            slide,
            bx,
            4.55,
            1.25,
            0.40,
            label,
            bg_hex=C_GOLD if active else "22486B",
            fg_hex="FFFFFF" if active else "94A3B8",
            size=11,
        )

    add_text(
        slide,
        1.05,
        5.25,
        5.5,
        1.1,
        "同行家人：父母 · 爱人 · 孩子（三代同堂全家福之旅）\n福利支持：公司全额承担旅游费用 · 珍藏全家的幸福记忆",
        size=12.5,
        color_hex="F8FAFC",
        line_spacing=1.45,
    )

    # 右侧封面主图（全家福）
    add_photo_frame(
        slide,
        pool,
        left=7.25,
        top=0.85,
        width=5.36,
        height=5.85,
        caption="三代同堂 · 艾多美 5 周年沪杭全家福留念",
        keywords=("全家", "合影", "家人", "父母", "外滩", "西湖", "cover", "01", "1"),
        theme="family",
    )


def build_slide_2_gratitude(prs: Presentation, pool: PhotoPool) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_header(
        slide,
        page_num="02 / 07",
        tag_text="感恩公司福利",
        title="五年耕耘，一份温暖全家的旅行礼遇",
        subtitle="在艾多美，每满 3、5、7、10 年都享有公司承担费用的员工旅游——让努力工作的成果，化作陪伴家人的幸福时光",
    )

    # 左侧：3/5/7/10 年福利时间轴卡片
    add_box(slide, 0.65, 1.35, 7.25, 2.55, fill_hex=C_CARD, border_hex=C_BORDER)
    add_text(
        slide,
        0.95,
        1.55,
        6.6,
        0.4,
        "艾多美员工旅游里程碑（费用由公司承担）",
        size=15,
        color_hex=C_NAVY,
        bold=True,
    )

    milestones = [
        ("满 3 年", "初绽芳华", "开启第一段感恩旅程", False),
        ("满 5 年", "五载同心", "携父母妻儿共游沪杭", True),
        ("满 7 年", "笃行致远", "期待下一次山海相逢", False),
        ("满 10 年", "十年荣耀", "长情陪伴与共同成长", False),
    ]
    for i, (yr, t1, t2, active) in enumerate(milestones):
        cx = 0.95 + i * 1.70
        cy = 2.08
        add_box(
            slide,
            cx,
            cy,
            1.55,
            1.55,
            fill_hex=C_GOLD_SOFT if active else C_BG,
            border_hex=C_GOLD if active else C_BORDER,
            border_pt=1.5 if active else 1.0,
        )
        add_pill(
            slide,
            cx + 0.20,
            cy + 0.14,
            1.15,
            0.32,
            yr,
            bg_hex=C_GOLD if active else C_ATOMY_SOFT,
            fg_hex="FFFFFF" if active else C_ATOMY,
            size=10.5,
        )
        add_text(
            slide,
            cx + 0.08,
            cy + 0.56,
            1.39,
            0.35,
            t1,
            size=13,
            color_hex=C_NAVY,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        add_text(
            slide,
            cx + 0.08,
            cy + 0.92,
            1.39,
            0.52,
            t2,
            size=10,
            color_hex=C_MUTED,
            align=PP_ALIGN.CENTER,
        )

    # 左下：本次行程概览卡片
    add_box(slide, 0.65, 4.12, 7.25, 2.85, fill_hex=C_CARD, border_hex=C_BORDER)
    add_text(
        slide,
        0.95,
        4.32,
        6.6,
        0.4,
        "我的 5 周年圆梦之旅：三代同堂游沪杭",
        size=15,
        color_hex=C_NAVY,
        bold=True,
    )
    summary_text = (
        "• 同行成员：带着父母、老婆和孩子，一家老小齐出游，其乐融融\n"
        "• 第一站 · 上海：漫步外滩、登金茂大厦 56 层享 300 元/位观景自助、游豫园与城隍庙、逛南京路\n"
        "• 第二站 · 杭州：泛舟西湖、登雷峰塔望湖光山色、全家共享龙虾与面包蟹海鲜大餐\n"
        "• 最深感触：平时忙于工作，借公司 5 周年旅游福利带父母和妻儿看风景、品美食，家人脸上的笑容就是最好的礼物"
    )
    add_text(
        slide,
        0.95,
        4.78,
        6.65,
        2.0,
        summary_text,
        size=12.5,
        color_hex=C_TEXT,
        line_spacing=1.42,
    )

    # 右侧：温馨出发/家人合影
    add_photo_frame(
        slide,
        pool,
        left=8.15,
        top=1.35,
        width=4.53,
        height=5.62,
        caption="带着父母、爱人与孩子欢喜出发",
        keywords=("出发", "父母", "孩子", "老婆", "家人", "高铁", "酒店", "02", "2"),
        theme="family",
    )


def build_slide_3_shanghai_bund_jinmao(prs: Presentation, pool: PhotoPool) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_header(
        slide,
        page_num="03 / 07",
        tag_text="上海站 · 云端与江景",
        title="外滩万国建筑群 & 金茂大厦 56 层云端观景自助",
        subtitle="吹着黄浦江畔的微风，在 56 层高空一边俯瞰魔都天际线，一边与家人举杯共享丰盛自助美食",
    )

    # 左侧信息卡 1：外滩
    add_box(slide, 0.65, 1.35, 4.65, 2.70, fill_hex=C_CARD, border_hex=C_BORDER)
    add_pill(slide, 0.92, 1.58, 1.55, 0.34, "打卡地标 · 外滩", bg_hex=C_ATOMY_SOFT, fg_hex=C_ATOMY, size=10.5)
    add_text(
        slide,
        0.92,
        2.02,
        4.1,
        0.4,
        "黄浦江畔的百年风华",
        size=16,
        color_hex=C_NAVY,
        bold=True,
    )
    add_text(
        slide,
        0.92,
        2.48,
        4.1,
        1.4,
        "陪着父母漫步外滩江堤，一边是庄重典雅的万国建筑博览群，对岸是陆家嘴高耸入云的摩天大楼。孩子兴奋地指着江面游船，一家人留下难忘合影。",
        size=12,
        color_hex=C_TEXT,
        line_spacing=1.4,
    )

    # 左侧信息卡 2：金茂大厦 56 层观景自助
    add_box(slide, 0.65, 4.25, 4.65, 2.72, fill_hex=C_GOLD_SOFT, border_hex=C_GOLD)
    add_pill(
        slide,
        0.92,
        4.48,
        2.65,
        0.34,
        "高空美食 · 300元/位观景自助",
        bg_hex=C_GOLD,
        fg_hex="FFFFFF",
        size=10.5,
    )
    add_text(
        slide,
        0.92,
        4.92,
        4.1,
        0.4,
        "金茂大厦 56 层云端盛宴",
        size=16,
        color_hex=C_NAVY,
        bold=True,
    )
    add_text(
        slide,
        0.92,
        5.38,
        4.1,
        1.45,
        "特地带全家登上金茂大厦 56 层，体验 300 元/位的高空观景自助餐。震撼的中庭挑高与窗外繁华景致相伴，父母和孩子品尝着中西式佳肴与精致甜品，赞不绝口。",
        size=12,
        color_hex=C_TEXT,
        line_spacing=1.4,
    )

    # 右侧两张照片并排
    add_photo_frame(
        slide,
        pool,
        left=5.55,
        top=1.35,
        width=3.45,
        height=5.62,
        caption="上海外滩｜江畔漫步与家人留影",
        keywords=("外滩", "东方明珠", "黄浦江", "上海", "bund", "03", "3"),
        theme="bund",
    )
    add_photo_frame(
        slide,
        pool,
        left=9.23,
        top=1.35,
        width=3.45,
        height=5.62,
        caption="金茂大厦56层｜300元/位云端观景自助",
        keywords=("金茂", "56", "自助", "餐厅", "jinmao", "04", "4"),
        theme="jinmao",
    )


def build_slide_4_shanghai_yuyuan_nanjing(prs: Presentation, pool: PhotoPool) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_header(
        slide,
        page_num="04 / 07",
        tag_text="上海站 · 海派风情",
        title="豫园江南亭台 · 城隍庙市井烟火 · 南京路繁华街景",
        subtitle="从曲径通幽的明代古典园林，到热闹非凡的老城厢小吃与中华商业第一街，感受古今交融的上海魅力",
    )

    cards = [
        (
            "豫园 & 城隍庙｜古典园林与地道风味",
            "九曲桥畔锦鲤成群，亭台楼阁飞檐翘角。陪父母细赏江南园林雅致，在城隍庙品尝南翔小笼包与传统糕点，烟火气十足。",
            ("豫园", "城隍庙", "九曲桥", "小笼", "园林", "05", "5"),
            "yuyuan",
        ),
        (
            "南京路步行街｜十里洋场的璀璨活力",
            "华灯初上时全家漫步南京路步行街，百年老字号与现代潮流交相辉映，老婆和孩子一路逛吃拍照，欢声笑语不断。",
            ("南京路", "步行街", "夜景", "街", "nanjing", "06", "6"),
            "nanjing",
        ),
    ]

    for idx, (cap, desc, kws, theme) in enumerate(cards):
        col_left = 0.65 + idx * 6.18
        add_photo_frame(
            slide,
            pool,
            left=col_left,
            top=1.35,
            width=5.85,
            height=4.20,
            caption=cap,
            keywords=kws,
            theme=theme,
        )
        add_box(slide, col_left, 5.70, 5.85, 1.28, fill_hex=C_CARD, border_hex=C_BORDER)
        add_text(
            slide,
            col_left + 0.22,
            5.82,
            5.40,
            1.05,
            desc,
            size=12,
            color_hex=C_TEXT,
            line_spacing=1.38,
        )


def build_slide_5_hangzhou_westlake(prs: Presentation, pool: PhotoPool) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_header(
        slide,
        page_num="05 / 07",
        tag_text="杭州站 · 诗画江南",
        title="西湖碧波荡漾 & 雷峰塔登高揽胜",
        subtitle="“欲把西湖比西子，淡妆浓抹总相宜”——放慢脚步，陪父母妻儿走入人间天堂的水墨画卷",
    )

    # 左侧两张大图
    add_photo_frame(
        slide,
        pool,
        left=0.65,
        top=1.35,
        width=3.85,
        height=5.62,
        caption="杭州西湖｜湖光山色与苏堤漫步",
        keywords=("西湖", "断桥", "游船", "三潭印月", "杭州", "westlake", "07", "7"),
        theme="westlake",
    )
    add_photo_frame(
        slide,
        pool,
        left=4.72,
        top=1.35,
        width=3.85,
        height=5.62,
        caption="雷峰塔｜登塔远眺西湖全景",
        keywords=("雷峰塔", "塔", "登高", "leifeng", "08", "8"),
        theme="leifeng",
    )

    # 右侧心得卡片
    add_box(slide, 8.80, 1.35, 3.88, 5.62, fill_hex=C_CARD, border_hex=C_BORDER)
    add_pill(
        slide,
        9.08,
        1.62,
        1.85,
        0.34,
        "西湖 · 雷峰塔漫记",
        bg_hex=C_ATOMY_SOFT,
        fg_hex=C_ATOMY,
        size=10.5,
    )
    add_text(
        slide,
        9.08,
        2.12,
        3.32,
        0.45,
        "山水之间，尽享天伦之乐",
        size=16,
        color_hex=C_NAVY,
        bold=True,
    )
    notes = (
        "• 漫步西湖畔：杨柳依依、碧波微澜，父母一路回忆着经典传说，步履轻松愉悦。\n\n"
        "• 登临雷峰塔：乘扶梯登塔远眺，整座西湖的亭台岛屿与杭州城景尽收眼底，孩子听着白娘子的故事格外入迷。\n\n"
        "• 慢节奏陪伴：不用赶时间的工作日，一家人并肩看风景，每一帧都值得珍藏。"
    )
    add_text(
        slide,
        9.08,
        2.68,
        3.32,
        4.0,
        notes,
        size=12,
        color_hex=C_TEXT,
        line_spacing=1.42,
    )


def build_slide_6_seafood_feast(prs: Presentation, pool: PhotoPool) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_header(
        slide,
        page_num="06 / 07",
        tag_text="舌尖上的狂欢",
        title="全家共享海鲜大餐：大龙虾、面包蟹满桌鲜香",
        subtitle="唯有爱与美食不可辜负——丰盛的龙虾与膏满肉肥的面包蟹，慰劳辛苦操劳的父母与家人",
    )

    # 左侧主图（大龙虾/海鲜全席）
    add_photo_frame(
        slide,
        pool,
        left=0.65,
        top=1.35,
        width=5.40,
        height=5.62,
        caption="海鲜盛宴｜鲜活大龙虾与丰盛海味上桌",
        keywords=("龙虾", "海鲜", "大餐", "美食", "lobster", "seafood", "09", "9"),
        theme="seafood",
    )

    # 中间副图（面包蟹/家人举杯）
    add_photo_frame(
        slide,
        pool,
        left=6.28,
        top=1.35,
        width=3.20,
        height=5.62,
        caption="膏满肉肥的面包蟹与家人欢聚",
        keywords=("面包蟹", "螃蟹", "蟹", "聚餐", "crab", "10"),
        theme="seafood",
    )

    # 右侧美食心得卡
    add_box(slide, 9.70, 1.35, 2.98, 5.62, fill_hex=C_GOLD_SOFT, border_hex=C_GOLD)
    add_pill(
        slide,
        9.95,
        1.62,
        1.85,
        0.34,
        "幸福味道 · 家的温度",
        bg_hex=C_GOLD,
        fg_hex="FFFFFF",
        size=10.5,
    )
    add_text(
        slide,
        9.95,
        2.12,
        2.48,
        0.45,
        "舌尖上的满足感",
        size=16,
        color_hex=C_NAVY,
        bold=True,
    )
    feast_text = (
        "• 招牌大龙虾：\n  虾肉紧实弹牙，鲜甜入味，孩子吃得津津有味。\n\n"
        "• 黄金面包蟹：\n  蟹壳饱满、蟹黄浓郁，特意盛给父母品尝。\n\n"
        "• 感恩时刻：\n  看着父母、老婆和孩子围坐一桌大快朵颐，深深感谢公司给予的旅游福利，让这份团圆更有滋有味。"
    )
    add_text(
        slide,
        9.95,
        2.68,
        2.48,
        4.0,
        feast_text,
        size=11.8,
        color_hex=C_TEXT,
        line_spacing=1.4,
    )


def build_slide_7_reflection(prs: Presentation, pool: PhotoPool) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_header(
        slide,
        page_num="07 / 07",
        tag_text="旅游心得与展望",
        title="五年感恩同行，满载幸福再出发",
        subtitle="一次旅行，温暖一个家庭；一份福利，凝聚一份归属感——迈向下一个 7 周年、10 周年",
    )

    reflections = [
        (
            "01  感恩艾多美 · 有温度的企业关怀",
            "每满 3、5、7、10 年的公费员工旅游，不仅是一份实在的物质激励，更是公司对员工背后家庭的体贴与敬意。在艾多美工作满 5 年，收获的不仅是职业成长，还有满满的幸福感与自豪感。",
        ),
        (
            "02  陪伴是最好的孝心与爱意",
            "带父母登金茂大厦 56 层看上海繁华、游西湖赏雷峰塔、吃龙虾与面包蟹海鲜大餐，看到父母欣慰的笑容、爱人放松的神情和孩子开阔眼界的雀跃，所有的奔波都化作了稳稳的幸福。",
        ),
        (
            "03  蓄力前行 · 奔赴 7 周年与 10 周年之约",
            "旅行归来，身心满电。家人的理解与支持是我最大的后盾，未来我将以更饱满的热情投入工作，与公司共同成长，期待满 7 年、满 10 年时再带家人去看看更广阔的世界！",
        ),
    ]

    for idx, (heading, body) in enumerate(reflections):
        top_y = 1.38 + idx * 1.88
        add_box(slide, 0.65, top_y, 7.35, 1.70, fill_hex=C_CARD, border_hex=C_BORDER)
        add_box(slide, 0.65, top_y, 0.12, 1.70, fill_hex=C_ATOMY if idx != 1 else C_GOLD, rounded=False)
        add_text(
            slide,
            0.95,
            top_y + 0.16,
            6.85,
            0.38,
            heading,
            size=14.5,
            color_hex=C_NAVY,
            bold=True,
        )
        add_text(
            slide,
            0.95,
            top_y + 0.56,
            6.85,
            1.02,
            body,
            size=11.8,
            color_hex=C_TEXT,
            line_spacing=1.38,
        )

    # 右侧圆满收官照
    add_photo_frame(
        slide,
        pool,
        left=8.25,
        top=1.38,
        width=4.43,
        height=5.46,
        caption="五年同行 · 心怀感恩 · 奔赴下一程山海",
        keywords=("全家", "返程", "合影", "西湖", "外滩", "end", "11"),
        theme="family",
    )


def generate_ppt(photo_dirs: Sequence[Path], output_path: Path) -> Tuple[int, Path]:
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    pool = PhotoPool(photo_dirs)

    build_slide_1_cover(prs, pool)
    build_slide_2_gratitude(prs, pool)
    build_slide_3_shanghai_bund_jinmao(prs, pool)
    build_slide_4_shanghai_yuyuan_nanjing(prs, pool)
    build_slide_5_hangzhou_westlake(prs, pool)
    build_slide_6_seafood_feast(prs, pool)
    build_slide_7_reflection(prs, pool)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return len(pool.photos), output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="生成《艾多美5周年员工旅游心得》PPT 演示文稿（自动配图排版）"
    )
    parser.add_argument(
        "-p",
        "--photos",
        default=DEFAULT_PHOTO_DIR,
        help=f"旅游照片所在文件夹路径（默认：{DEFAULT_PHOTO_DIR}）",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=DEFAULT_OUTPUT,
        help=f"输出的 PPTX 文件路径（默认：{DEFAULT_OUTPUT}）",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    candidate_dirs = [
        Path(args.photos).expanduser(),
        Path(__file__).resolve().parent / "photos",
        Path(__file__).resolve().parent / "5周年旅游照片",
    ]
    count, out = generate_ppt(candidate_dirs, Path(args.output).expanduser())
    if count > 0:
        print(f"已从照片目录载入 {count} 张实拍照片，成功生成 PPT：{out.resolve()}")
    else:
        print(
            f"未在当前环境中检测到本地目录 '{args.photos}' 下的照片，"
            f"已使用内置沪杭主题风景占位图生成 PPT：{out.resolve()}\n"
            f"提示：在您本机运行 `python3 make_travel_ppt.py` 即可自动读取 "
            f"'{DEFAULT_PHOTO_DIR}' 中的实拍照片并完成裁剪配图。"
        )


if __name__ == "__main__":
    main()
