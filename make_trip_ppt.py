#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
艾多美 5 周年员工旅游心得 PPT 生成器
====================================

自动读取本地照片目录（默认：/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片，
或当前目录下的 photos/ 文件夹），将照片智能匹配或按顺序排入页面，生成 6 页 16:9 宽屏 PPT。
若本地目录中暂未扫描到照片，脚本会自动生成带位置提示的精美占位框，
您也可以打开生成的 PPTX 直接右键“更改图片”替换为您心仪的照片。

用法示例：
  # 直接运行（自动查找默认路径 /Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片）
  python3 make_trip_ppt.py

  # 指定照片文件夹路径和输出文件名
  python3 make_trip_ppt.py -p "/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片" -o 艾多美5周年员工旅游心得.pptx
"""

import argparse
import io
import os
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from PIL import Image, ImageDraw, ImageFont, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

DEFAULT_MAC_PHOTO_DIR = "/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片"
FALLBACK_LOCAL_DIR = "photos"
DEFAULT_OUTPUT = "艾多美5周年员工旅游心得.pptx"

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".tif"}

# 配色方案：艾多美蓝 + 温暖米白 + 柔和金棕
ATOMY_BLUE = "00A0E9"
ATOMY_DARK_BLUE = "004B87"
DEEP_NAVY = "0F2537"
WARM_GOLD = "C88A32"
GOLD_SOFT = "FDF6EC"
BLUE_SOFT = "EBF6FC"
BG_CREAM = "F8F7F4"
WHITE = "FFFFFF"
INK = "1C242B"
BODY = "3A4652"
MUTED = "73808C"
BORDER = "E2E6EA"

FONT_SANS = "PingFang SC"
FONT_FALLBACK = "Microsoft YaHei"

SLIDE_W = 13.333333
SLIDE_H = 7.5

# 9 个照片槽位的定义：(槽位key, 标题说明, 匹配关键词列表, 推荐比例宽, 推荐比例高)
PHOTO_SLOTS = [
    (
        "cover_family",
        "全家福 / 旅途合影",
        ["全家", "合影", "家人", "父母", "老婆", "孩子", "家庭", "出发", "cover", "family", "1"],
        800,
        960,
    ),
    (
        "bund",
        "上海 · 外滩江景与万国建筑",
        ["外滩", "东方明珠", "陆家嘴", "黄浦江", "江景", "夜景", "上海", "bund"],
        900,
        640,
    ),
    (
        "jinmao",
        "金茂大厦56层 · 观景自助餐",
        ["金茂", "56", "自助", "云端", "观景", "金茂大厦", "餐厅", "jinmao", "buffet"],
        900,
        640,
    ),
    (
        "yuyuan",
        "上海 · 豫园 & 城隍庙古韵",
        ["豫园", "城隍庙", "九曲桥", "古镇", "园林", "yuyuan"],
        800,
        600,
    ),
    (
        "nanjing",
        "上海 · 南京路步行街",
        ["南京路", "步行街", "南京东路", "街景", "和平饭店", "nanjing"],
        800,
        600,
    ),
    (
        "westlake",
        "杭州 · 西湖漫步",
        ["西湖", "断桥", "苏堤", "游船", "三潭印月", "湖滨", "杭州", "westlake", "xihu"],
        800,
        600,
    ),
    (
        "leifeng",
        "杭州 · 雷峰塔登高望远",
        ["雷峰塔", "雷峰", "夕照", "登塔", "leifeng"],
        800,
        600,
    ),
    (
        "seafood_1",
        "海鲜大餐 · 龙虾与面包蟹",
        ["海鲜", "龙虾", "面包蟹", "螃蟹", "大餐", "美食", "吃", "seafood", "lobster", "crab"],
        900,
        620,
    ),
    (
        "seafood_2",
        "家人欢聚享美食 / 旅途温馨瞬间",
        ["聚餐", "晚餐", "美食2", "海鲜2", "开心", "笑容", "干杯", "dinner", "happy"],
        900,
        620,
    ),
]


def rgb(hex_str: str) -> RGBColor:
    return RGBColor.from_string(hex_str.lstrip("#"))


def _set_text_frame_margins(tf, valign="top", margin_pt=0):
    body = tf._txBody.find(qn("a:bodyPr"))
    body.set("anchor", {"top": "t", "middle": "ctr", "bottom": "b"}[valign])
    ins = str(int(Pt(margin_pt)))
    for key in ("lIns", "rIns", "tIns", "bIns"):
        body.set(key, ins)


def _add_run(p, text: str, size: float, color: str, bold: bool = False, font: str = FONT_SANS):
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    run.font.name = font
    rpr = run._r.get_or_add_rPr()
    rpr.set("lang", "zh-CN")
    rpr.set("altLang", "en-US")
    for tag in ("latin", "ea", "cs"):
        el = rpr.find(qn(f"a:{tag}"))
        if el is None:
            el = rpr.makeelement(qn(f"a:{tag}"), {"typeface": font})
            rpr.append(el)
        else:
            el.set("typeface", font)


def add_textbox(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    paragraphs: List[dict],
    valign: str = "top",
    margin_pt: float = 0,
):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    _set_text_frame_margins(tf, valign=valign, margin_pt=margin_pt)
    align_map = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
    }
    for idx, info in enumerate(paragraphs):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align_map.get(info.get("align", "left"), PP_ALIGN.LEFT)
        if "line" in info:
            p.line_spacing = info["line"]
        if info.get("before"):
            p.space_before = Pt(info["before"])
        if info.get("after"):
            p.space_after = Pt(info["after"])
        runs = info.get("runs")
        if runs:
            for r in runs:
                _add_run(
                    p,
                    r["text"],
                    r.get("size", 14),
                    r.get("color", BODY),
                    r.get("bold", False),
                    r.get("font", FONT_SANS),
                )
        else:
            _add_run(
                p,
                info.get("text", ""),
                info.get("size", 14),
                info.get("color", BODY),
                info.get("bold", False),
                info.get("font", FONT_SANS),
            )
    return shape


def add_rect(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    fill_color: str,
    line_color: Optional[str] = None,
    line_width_pt: float = 1.0,
    rounded: bool = False,
):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill_color)
    if rounded and shape.adjustments:
        shape.adjustments[0] = 0.03
    if line_color:
        shape.line.color.rgb = rgb(line_color)
        shape.line.width = Pt(line_width_pt)
    else:
        shape.line.fill.background()
    return shape


def set_slide_bg(slide, color: str):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color)


def add_header(slide, tag_text: str, title_text: str, subtitle_text: str, page_num: int, total_pages: int = 6):
    """统一的顶部页眉与页码设计。"""
    # 顶部细装饰条
    add_rect(slide, 0, 0, SLIDE_W, 0.07, ATOMY_BLUE)

    # 胶囊标签
    add_rect(slide, 0.65, 0.36, 2.35, 0.34, BLUE_SOFT, rounded=True)
    add_textbox(
        slide,
        0.65,
        0.36,
        2.35,
        0.34,
        [{"text": tag_text, "size": 11, "color": ATOMY_DARK_BLUE, "bold": True, "align": "center"}],
        valign="middle",
    )

    # 右上角公司福利标识与页码
    add_textbox(
        slide,
        8.5,
        0.38,
        4.18,
        0.32,
        [
            {
                "align": "right",
                "runs": [
                    {"text": "ATOMY 艾多美 · 五周年员工福利旅游   ", "size": 10.5, "color": MUTED},
                    {"text": f"0{page_num} / 0{total_pages}", "size": 11, "color": ATOMY_BLUE, "bold": True},
                ],
            }
        ],
        valign="middle",
    )

    # 主标题与副标题
    add_textbox(
        slide,
        0.65,
        0.78,
        12.0,
        0.85,
        [
            {"text": title_text, "size": 24, "color": INK, "bold": True, "after": 4},
            {"text": subtitle_text, "size": 13, "color": MUTED},
        ],
    )


def discover_photos(photo_dirs: List[str]) -> Tuple[List[Path], Optional[Path]]:
    """扫描传入的照片目录列表，返回找到的所有图片路径及实际命中的目录。"""
    for d in photo_dirs:
        if not d:
            continue
        p = Path(d).expanduser()
        if p.exists() and p.is_dir():
            files = sorted(
                [f for f in p.iterdir() if f.is_file() and f.suffix.lower() in IMAGE_EXTS],
                key=lambda x: x.name.lower(),
            )
            if files:
                return files, p
    return [], None


def assign_photos_to_slots(files: List[Path]) -> Dict[str, Optional[Path]]:
    """
    先按文件名关键词匹配各景点槽位，剩余未匹配槽位按照片顺序自动填充；
    如果照片数量少于槽位数，则循环使用已有照片或保留为 None（生成占位图）。
    """
    assigned: Dict[str, Optional[Path]] = {slot[0]: None for slot in PHOTO_SLOTS}
    if not files:
        return assigned

    used_files = set()

    # 第一轮：按文件名包含关键词精确匹配
    for slot_key, _label, keywords, _w, _h in PHOTO_SLOTS:
        for f in files:
            if f in used_files:
                continue
            stem_lower = f.stem.lower()
            if any(kw.lower() in stem_lower for kw in keywords):
                assigned[slot_key] = f
                used_files.add(f)
                break

    # 第二轮：剩余照片按顺序填入未分配的槽位
    remaining = [f for f in files if f not in used_files]
    rem_idx = 0
    for slot_key, _label, _keywords, _w, _h in PHOTO_SLOTS:
        if assigned[slot_key] is None and rem_idx < len(remaining):
            assigned[slot_key] = remaining[rem_idx]
            rem_idx += 1

    # 第三轮：如果还有空槽位且至少有一张照片，循环复用照片保证每页都有真实照片展示
    for idx, (slot_key, _label, _keywords, _w, _h) in enumerate(PHOTO_SLOTS):
        if assigned[slot_key] is None and files:
            assigned[slot_key] = files[idx % len(files)]

    return assigned


def _load_cjk_font(size: int):
    font_candidates = [
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
        "C:/Windows/Fonts/msyh.ttc",
    ]
    for fp in font_candidates:
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                continue
    return ImageFont.load_default()


def prepare_image_stream(
    photo_path: Optional[Path],
    target_w_px: int,
    target_h_px: int,
    title_label: str,
    theme_hint: str = "blue",
) -> io.BytesIO:
    """
    将真实照片按目标宽高比居中裁剪（修正 EXIF 方向，不变形）；
    若未找到本地照片，则生成一张优雅的渐变风景占位图片，方便直接在 PPT 中展示或右键替换。
    """
    if photo_path and photo_path.exists():
        try:
            with Image.open(photo_path) as im:
                im = ImageOps.exif_transpose(im)
                if im.mode not in ("RGB", "RGBA"):
                    im = im.convert("RGB")
                fitted = ImageOps.fit(
                    im.convert("RGB"),
                    (target_w_px, target_h_px),
                    method=Image.Resampling.LANCZOS,
                    centering=(0.5, 0.5),
                )
                buf = io.BytesIO()
                fitted.save(buf, format="JPEG", quality=92)
                buf.seek(0)
                return buf
        except Exception as exc:
            print(f"[提示] 读取照片 {photo_path} 失败 ({exc})，将使用占位图。")

    # 生成高颜值占位图
    img = Image.new("RGB", (target_w_px, target_h_px))
    draw = ImageDraw.Draw(img)

    palettes = {
        "blue": ((18, 53, 88), (42, 118, 172), (210, 232, 248)),
        "gold": ((74, 48, 24), (168, 112, 48), (248, 232, 206)),
        "lake": ((22, 68, 76), (56, 132, 136), (214, 238, 236)),
        "warm": ((82, 38, 38), (176, 84, 68), (250, 226, 218)),
    }
    top_c, bot_c, accent_c = palettes.get(theme_hint, palettes["blue"])

    for y in range(target_h_px):
        ratio = y / max(target_h_px - 1, 1)
        r = int(top_c[0] * (1 - ratio) + bot_c[0] * ratio)
        g = int(top_c[1] * (1 - ratio) + bot_c[1] * ratio)
        b = int(top_c[2] * (1 - ratio) + bot_c[2] * ratio)
        draw.line([(0, y), (target_w_px, y)], fill=(r, g, b))

    # 绘制简约装饰几何线条（山峦/天际线意象）
    w, h = target_w_px, target_h_px
    draw.ellipse([int(w * 0.72), int(h * 0.14), int(w * 0.86), int(h * 0.14 + w * 0.14)], fill=accent_c)
    draw.polygon(
        [(0, h), (0, int(h * 0.72)), (int(w * 0.32), int(h * 0.54)), (int(w * 0.68), int(h * 0.76)), (w, int(h * 0.58)), (w, h)],
        fill=(top_c[0], top_c[1], top_c[2]),
    )

    # 内边框
    pad = 22
    draw.rectangle([pad, pad, w - pad, h - pad], outline=accent_c, width=2)

    font_title = _load_cjk_font(max(24, int(min(w, h) * 0.052)))
    font_sub = _load_cjk_font(max(16, int(min(w, h) * 0.032)))

    text_main = f"【照片位置】{title_label}"
    text_sub = "运行脚本自动载入本地照片，或在 PPT 中右键 →「更改图片」"

    bbox1 = draw.textbbox((0, 0), text_main, font=font_title)
    tw1, th1 = bbox1[2] - bbox1[0], bbox1[3] - bbox1[1]
    draw.text(((w - tw1) / 2, h * 0.40 - th1 / 2), text_main, fill=(255, 255, 255), font=font_title)

    bbox2 = draw.textbbox((0, 0), text_sub, font=font_sub)
    tw2, th2 = bbox2[2] - bbox2[0], bbox2[3] - bbox2[1]
    draw.text(((w - tw2) / 2, h * 0.52 - th2 / 2), text_sub, fill=accent_c, font=font_sub)

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=90)
    buf.seek(0)
    return buf


def add_photo_card(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    photo_path: Optional[Path],
    caption_title: str,
    caption_desc: str,
    theme_hint: str = "blue",
):
    """绘制带白色拍立得相框和底部图注的照片卡片。"""
    add_rect(slide, x, y, w, h, WHITE, line_color=BORDER, line_width_pt=1.0, rounded=True)

    pad = 0.12
    caption_h = 0.76
    img_x = x + pad
    img_y = y + pad
    img_w = w - pad * 2
    img_h = h - pad * 2 - caption_h

    px_w = max(400, int(img_w * 180))
    px_h = max(300, int(img_h * 180))
    stream = prepare_image_stream(photo_path, px_w, px_h, caption_title, theme_hint=theme_hint)
    slide.shapes.add_picture(stream, Inches(img_x), Inches(img_y), width=Inches(img_w), height=Inches(img_h))

    add_textbox(
        slide,
        img_x + 0.04,
        img_y + img_h + 0.08,
        img_w - 0.08,
        caption_h - 0.06,
        [
            {"text": caption_title, "size": 13, "color": INK, "bold": True, "after": 2},
            {"text": caption_desc, "size": 10.5, "color": MUTED, "line": 1.2},
        ],
    )


def build_slide_1_cover(prs, slots: Dict[str, Optional[Path]]):
    """第 1 页：封面 —— 感恩艾多美 · 五周年华东亲情之旅。"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, DEEP_NAVY)

    # 顶部品牌色条与底部暖金装饰条
    add_rect(slide, 0, 0, SLIDE_W, 0.10, ATOMY_BLUE)
    add_rect(slide, 0, SLIDE_H - 0.12, SLIDE_W, 0.12, WARM_GOLD)

    # 左侧内容区
    add_rect(slide, 0.78, 0.78, 3.30, 0.40, ATOMY_DARK_BLUE, rounded=True)
    add_textbox(
        slide,
        0.78,
        0.78,
        3.30,
        0.40,
        [{"text": "ATOMY 艾多美 · 5周年员工福利旅游", "size": 12, "color": WHITE, "bold": True, "align": "center"}],
        valign="middle",
    )

    add_textbox(
        slide,
        0.78,
        1.42,
        6.9,
        2.35,
        [
            {"text": "五年相伴 · 感恩同行", "size": 40, "color": WHITE, "bold": True, "after": 10},
            {"text": "携父母妻儿共赴上海 & 杭州幸福之旅", "size": 22, "color": "7FD4FF", "bold": True, "after": 16},
            {
                "text": "在艾多美，每满 3、5、7、10 年都有一次由公司承担费用的专属旅游福利。\n转眼入职满五周年，最好的礼物莫过于带着父母、老婆和孩子一起看世界。",
                "size": 13.5,
                "color": "D4E2EE",
                "line": 1.45,
            },
        ],
    )

    # 三个核心信息小卡片
    highlights = [
        ("公司福利机制", "满 3 / 5 / 7 / 10 年\n专享带薪员工旅游"),
        ("同行幸福家人", "父母 · 老婆 · 孩子\n三代同堂温馨出游"),
        ("双城核心行程", "上海外滩 · 金茂56层\n豫园南京路 · 杭州西湖"),
    ]
    for i, (h_title, h_desc) in enumerate(highlights):
        cx = 0.78 + i * 2.28
        cy = 4.25
        add_rect(slide, cx, cy, 2.12, 1.45, "17364F", line_color="2A557A", rounded=True)
        add_rect(slide, cx, cy, 2.12, 0.06, ATOMY_BLUE if i < 2 else WARM_GOLD)
        add_textbox(
            slide,
            cx + 0.16,
            cy + 0.18,
            1.80,
            1.15,
            [
                {"text": h_title, "size": 13, "color": "7FD4FF" if i < 2 else "F5C478", "bold": True, "after": 6},
                {"text": h_desc, "size": 11.5, "color": WHITE, "line": 1.35},
            ],
        )

    # 底部寄语
    add_textbox(
        slide,
        0.78,
        6.18,
        6.8,
        0.65,
        [
            {
                "text": "“工作的意义不仅是创造价值，更是让家人共享成长的喜悦。感谢艾多美，让这份5周年的仪式感如此圆满。”",
                "size": 12,
                "color": "A9C1D6",
            }
        ],
    )

    # 右侧大图相框
    add_photo_card(
        slide,
        8.05,
        0.78,
        4.55,
        5.92,
        slots.get("cover_family"),
        "五周年 · 全家幸福合影",
        "带着父母、老婆和孩子，开启上海与杭州的双城美好时光",
        theme_hint="blue",
    )


def build_slide_2_welfare_overview(prs):
    """第 2 页：福利制度与行程总览 —— 感恩公司关怀，三代同堂出发。"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG_CREAM)
    add_header(
        slide,
        "五周年福利 · 行程总览",
        "满五周年专属礼遇：公司买单，带全家去旅行",
        "艾多美 3 / 5 / 7 / 10 年员工旅游制度，让努力工作的我们能够回馈家人陪伴",
        page_num=2,
    )

    # 上半部分：艾多美周年旅游里程碑时间轴
    add_rect(slide, 0.65, 1.78, 12.03, 2.05, WHITE, line_color=BORDER, rounded=True)
    add_textbox(
        slide,
        0.95,
        1.94,
        11.4,
        0.4,
        [
            {
                "runs": [
                    {"text": "艾多美员工旅游福利阶梯  ", "size": 15, "color": INK, "bold": True},
                    {"text": "｜ 每一步成长都有公司的温暖见证（费用由公司承担）", "size": 12, "color": MUTED},
                ]
            }
        ],
    )

    milestones = [
        ("满 3 周年", "初绽光芒 · 蓄力前行", False),
        ("满 5 周年（本次）", "携父母妻儿 · 沪杭双城游", True),
        ("满 7 周年", "深耕蜕变 · 再启新程", False),
        ("满 10 周年", "十年同心 · 荣耀远航", False),
    ]
    # 连接横线
    add_rect(slide, 1.5, 2.95, 10.3, 0.04, BORDER)

    for idx, (m_year, m_desc, is_current) in enumerate(milestones):
        mx = 0.95 + idx * 2.92
        my = 2.48
        card_fill = GOLD_SOFT if is_current else "F3F6F9"
        border_c = WARM_GOLD if is_current else BORDER
        badge_c = WARM_GOLD if is_current else ATOMY_BLUE
        add_rect(slide, mx, my, 2.68, 1.12, card_fill, line_color=border_c, line_width_pt=1.5 if is_current else 1.0, rounded=True)
        add_rect(slide, mx + 0.16, my + 0.16, 2.36, 0.34, badge_c, rounded=True)
        add_textbox(
            slide,
            mx + 0.16,
            my + 0.16,
            2.36,
            0.34,
            [{"text": m_year, "size": 11.5, "color": WHITE, "bold": True, "align": "center"}],
            valign="middle",
        )
        add_textbox(
            slide,
            mx + 0.14,
            my + 0.58,
            2.40,
            0.42,
            [{"text": m_desc, "size": 11.5, "color": INK if is_current else BODY, "bold": is_current, "align": "center"}],
            valign="middle",
        )

    # 下半部分：左中右三张详情卡（同行家人、上海站、杭州站&美食）
    cards = [
        (
            "01  温暖同行 · 三代同堂",
            ATOMY_BLUE,
            "与最爱的人共享公司福利",
            [
                "• 同行成员：父母、老婆、孩子与我",
                "• 平时忙于工作，难得带爸妈和妻儿远行",
                "• 老人看风景舒心、孩子长见识开心",
                "• 费用由公司承担，全家零负担尽情体验",
            ],
        ),
        (
            "02  第一站 · 魔都上海",
            ATOMY_DARK_BLUE,
            "繁华天际线与海派老城厢",
            [
                "• 外滩漫步：饱览黄浦江两岸万国建筑群",
                "• 金茂大厦56层：300元/位云端观景自助餐",
                "• 豫园 & 城隍庙：九曲桥畔感受江南古韵",
                "• 南京路步行街：打卡十里洋场热闹夜景",
            ],
        ),
        (
            "03  第二站 · 诗画杭州 & 大餐",
            WARM_GOLD,
            "西子湖光山色与海鲜饕餮盛宴",
            [
                "• 西湖泛舟漫步：看湖光潋滟、柳浪闻莺",
                "• 雷峰塔登高：俯瞰西湖全景，品千年传说",
                "• 海鲜大餐：大龙虾、面包蟹等鲜美丰盛",
                "• 一路欢声笑语，留下全家珍贵回忆",
            ],
        ),
    ]

    for i, (c_tag, c_color, c_title, bullets) in enumerate(cards):
        cx = 0.65 + i * 4.10
        cy = 4.08
        cw = 3.83
        ch = 2.92
        add_rect(slide, cx, cy, cw, ch, WHITE, line_color=BORDER, rounded=True)
        add_rect(slide, cx, cy, cw, 0.08, c_color)
        add_textbox(
            slide,
            cx + 0.24,
            cy + 0.22,
            cw - 0.48,
            0.68,
            [
                {"text": c_tag, "size": 11.5, "color": c_color, "bold": True, "after": 4},
                {"text": c_title, "size": 16, "color": INK, "bold": True},
            ],
        )
        bullet_paras = [{"text": b, "size": 12, "color": BODY, "after": 6, "line": 1.3} for b in bullets]
        add_textbox(slide, cx + 0.24, cy + 1.02, cw - 0.48, 1.72, bullet_paras)


def build_slide_3_shanghai_skyline(prs, slots: Dict[str, Optional[Path]]):
    """第 3 页：上海站（上） —— 外滩江风与金茂大厦 56 层云端观景自助。"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG_CREAM)
    add_header(
        slide,
        "上海站 · 云端与江景",
        "外滩万国建筑群 & 金茂大厦56层观景自助",
        "吹着黄浦江畔的晚风，在56层高空边俯瞰魔都天际线边享300元/位自助美食",
        page_num=3,
    )

    # 左上：外滩大图
    add_photo_card(
        slide,
        0.65,
        1.78,
        5.80,
        5.22,
        slots.get("bund"),
        "上海外滩 · 浦江两岸风光",
        "带父母和妻儿漫步外滩，身后是陆家嘴天际线与百年万国建筑博览群",
        theme_hint="blue",
    )

    # 右上：金茂大厦56层照片
    add_photo_card(
        slide,
        6.68,
        1.78,
        6.00,
        3.45,
        slots.get("jinmao"),
        "金茂大厦 56 层 · 300元/位高空观景自助餐",
        "置身金茂君悦56层云端中庭，丰盛料理配震撼高空城景，爸妈和孩子都赞不绝口",
        theme_hint="gold",
    )

    # 右下：心得感悟卡片
    add_rect(slide, 6.68, 5.40, 6.00, 1.60, WHITE, line_color=BORDER, rounded=True)
    add_rect(slide, 6.68, 5.40, 0.10, 1.60, WARM_GOLD)
    add_textbox(
        slide,
        6.96,
        5.54,
        5.55,
        1.35,
        [
            {"text": "旅途高光时刻 · 云端家宴", "size": 14, "color": WARM_GOLD, "bold": True, "after": 4},
            {
                "text": "特意带家人体验了金茂大厦56层（300元/位）的观景自助。看着父母在窗前欣喜地俯瞰上海风光、老婆和孩子开心地挑选美食，心里满是幸福——感谢公司报销旅游费用，让我们尽情享受高品质的家庭时光。",
                "size": 11.5,
                "color": BODY,
                "line": 1.38,
            },
        ],
    )


def build_slide_4_shanghai_culture(prs, slots: Dict[str, Optional[Path]]):
    """第 4 页：上海站（下） —— 豫园、城隍庙与南京路步行街。"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG_CREAM)
    add_header(
        slide,
        "上海站 · 海派风情",
        "豫园城隍庙古韵 & 南京路十里洋场",
        "一步入江南古典园林，一步踏入中华商业第一街，感受老上海与新繁华的交融",
        page_num=4,
    )

    # 左侧：豫园 & 城隍庙照片卡
    add_photo_card(
        slide,
        0.65,
        1.78,
        4.35,
        5.22,
        slots.get("yuyuan"),
        "豫园 & 城隍庙 · 亭台楼阁九曲桥",
        "飞檐翘角、小桥流水，带父母品尝地道南翔小笼与海派小吃",
        theme_hint="warm",
    )

    # 中间：南京路照片卡
    add_photo_card(
        slide,
        5.20,
        1.78,
        4.35,
        5.22,
        slots.get("nanjing"),
        "南京路步行街 · 霓虹璀璨百年商街",
        "乘铛铛车、逛老字号与潮流地标，孩子一路欢呼雀跃",
        theme_hint="blue",
    )

    # 右侧：海派游览札记栏
    add_rect(slide, 9.75, 1.78, 2.93, 5.22, WHITE, line_color=BORDER, rounded=True)
    add_rect(slide, 9.75, 1.78, 2.93, 0.08, ATOMY_BLUE)
    add_textbox(
        slide,
        9.98,
        2.02,
        2.48,
        4.80,
        [
            {"text": "海派漫步札记", "size": 16, "color": INK, "bold": True, "after": 12},
            {"text": "▪ 豫园雅韵", "size": 13, "color": ATOMY_DARK_BLUE, "bold": True, "after": 4},
            {
                "text": "江南古典园林一步一景，父母最喜欢这里的传统建筑与慢节奏韵味。",
                "size": 11.5,
                "color": BODY,
                "line": 1.35,
                "after": 12,
            },
            {"text": "▪ 城隍庙烟火气", "size": 13, "color": ATOMY_DARK_BLUE, "bold": True, "after": 4},
            {
                "text": "九曲桥畔热闹非凡，热气腾腾的特色点心让全家人胃口大开。",
                "size": 11.5,
                "color": BODY,
                "line": 1.35,
                "after": 12,
            },
            {"text": "▪ 南京路繁华夜色", "size": 13, "color": WARM_GOLD, "bold": True, "after": 4},
            {
                "text": "从南京路一直走到外滩，灯火辉煌的街景定格了全家人的灿烂笑容。",
                "size": 11.5,
                "color": BODY,
                "line": 1.35,
            },
        ],
    )


def build_slide_5_hangzhou(prs, slots: Dict[str, Optional[Path]]):
    """第 5 页：杭州站 —— 西湖山水画卷与雷峰塔登高。"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG_CREAM)
    add_header(
        slide,
        "杭州站 · 人间天堂",
        "漫步西子湖畔 & 登临雷峰塔俯瞰全湖",
        "告别魔都的摩登繁华，转赴诗画江南，在西湖烟雨与雷峰夕照中慢享天伦之乐",
        page_num=5,
    )

    # 左侧：西湖与雷峰塔行程体验面板
    add_rect(slide, 0.65, 1.78, 3.65, 5.22, WHITE, line_color=BORDER, rounded=True)
    add_rect(slide, 0.65, 1.78, 3.65, 0.08, "2A7B76")
    add_textbox(
        slide,
        0.92,
        2.02,
        3.12,
        4.80,
        [
            {"text": "江南诗意慢时光", "size": 17, "color": INK, "bold": True, "after": 12},
            {
                "text": "“欲把西湖比西子，淡妆浓抹总相宜。”到了杭州，脚步自然慢了下来。",
                "size": 12,
                "color": "2A7B76",
                "bold": True,
                "line": 1.35,
                "after": 14,
            },
            {"text": "01  西湖湖畔漫步", "size": 13, "color": INK, "bold": True, "after": 4},
            {
                "text": "微风拂过湖面，杨柳依依。陪着爸妈和老婆孩子沿湖慢行，随手一拍都是如画风景。",
                "size": 11.5,
                "color": BODY,
                "line": 1.35,
                "after": 12,
            },
            {"text": "02  雷峰塔登高望远", "size": 13, "color": INK, "bold": True, "after": 4},
            {
                "text": "乘扶梯登上雷峰塔顶，整片西湖的碧波与远山尽收眼底；给孩子讲白娘子的传说，寓教于游。",
                "size": 11.5,
                "color": BODY,
                "line": 1.35,
                "after": 14,
            },
            {
                "text": "✦ 旅行感悟：最美的风景在路上，更在身边家人的笑脸里。",
                "size": 11.5,
                "color": WARM_GOLD,
                "bold": True,
                "line": 1.35,
            },
        ],
    )

    # 中间：西湖照片卡
    add_photo_card(
        slide,
        4.52,
        1.78,
        4.00,
        5.22,
        slots.get("westlake"),
        "杭州西湖 · 碧波荡漾画中游",
        "一家人漫步西湖岸边，看亭台映水、游船点点，惬意又治愈",
        theme_hint="lake",
    )

    # 右侧：雷峰塔照片卡
    add_photo_card(
        slide,
        8.70,
        1.78,
        3.98,
        5.22,
        slots.get("leifeng"),
        "雷峰塔 · 登高凭栏揽西湖全景",
        "古塔巍峨，登顶远眺苏堤与三潭印月，共赏西湖千年人文底蕴",
        theme_hint="gold",
    )


def build_slide_6_seafood_and_reflection(prs, slots: Dict[str, Optional[Path]]):
    """第 6 页：舌尖盛宴与五周年感恩心得 —— 海鲜大餐（龙虾、面包蟹）与未来展望。"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG_CREAM)
    add_header(
        slide,
        "舌尖盛宴 · 五年心得",
        "龙虾面包蟹海鲜大餐 & 五周年感恩寄语",
        "用一顿丰盛的海鲜大餐犒赏全家人的味蕾，把公司的温暖化作继续前行的动力",
        page_num=6,
    )

    # 左侧两张美食/欢聚照片
    add_photo_card(
        slide,
        0.65,
        1.78,
        3.85,
        5.22,
        slots.get("seafood_1"),
        "海鲜大餐 · 大龙虾 & 面包蟹",
        "膏满肉肥的面包蟹、鲜甜弹牙的大龙虾，全家人大快朵颐",
        theme_hint="warm",
    )

    add_photo_card(
        slide,
        4.68,
        1.78,
        3.85,
        5.22,
        slots.get("seafood_2"),
        "举杯欢聚 · 共享幸福滋味",
        "美食配欢笑，父母吃得开心、老婆孩子满脸幸福",
        theme_hint="gold",
    )

    # 右侧：五周年旅游心得总结深色卡
    add_rect(slide, 8.72, 1.78, 3.96, 5.22, DEEP_NAVY, rounded=True)
    add_rect(slide, 8.72, 1.78, 3.96, 0.08, WARM_GOLD)

    add_textbox(
        slide,
        9.00,
        2.02,
        3.40,
        4.75,
        [
            {"text": "五周年 · 旅游心得感悟", "size": 17, "color": WHITE, "bold": True, "after": 12},
            {"text": "1. 感恩艾多美的人文关怀", "size": 13, "color": "7FD4FF", "bold": True, "after": 4},
            {
                "text": "满3、5、7、10年的公费旅游制度，不仅是一份福利，更是公司对员工付出与家庭支持的深切认可。",
                "size": 11.2,
                "color": "D8E6F2",
                "line": 1.35,
                "after": 12,
            },
            {"text": "2. 给家人最好的陪伴礼", "size": 13, "color": "F5C478", "bold": True, "after": 4},
            {
                "text": "从上海外滩、金茂56层自助、豫园城隍庙、南京路，到杭州西湖、雷峰塔和海鲜大餐，看到父母妻儿的笑容，倍感自豪。",
                "size": 11.2,
                "color": "D8E6F2",
                "line": 1.35,
                "after": 12,
            },
            {"text": "3. 满载能量，奔赴下一个里程碑", "size": 13, "color": "7FD4FF", "bold": True, "after": 4},
            {
                "text": "五年是节点更是新起点。未来将继续怀揣感恩之心努力工作，期待满 7 年、10 年时再带家人出发！",
                "size": 11.2,
                "color": "D8E6F2",
                "line": 1.35,
            },
        ],
    )


def generate_ppt(photo_dir: Optional[str], output_path: str) -> Path:
    candidate_dirs = []
    if photo_dir:
        candidate_dirs.append(photo_dir)
    candidate_dirs.extend([DEFAULT_MAC_PHOTO_DIR, FALLBACK_LOCAL_DIR])

    files, matched_dir = discover_photos(candidate_dirs)
    if matched_dir:
        print(f"[信息] 已从目录加载 {len(files)} 张照片：{matched_dir}")
    else:
        print(
            f"[提示] 当前环境未检测到本地照片目录（{DEFAULT_MAC_PHOTO_DIR} 或 ./{FALLBACK_LOCAL_DIR}）。\n"
            "       已自动生成带景点名称的高清占位框。在本机运行本脚本即可自动嵌入真实照片，\n"
            "       或直接在 PowerPoint / Keynote / WPS 中右键图片选择「更改图片」替换。"
        )

    slots = assign_photos_to_slots(files)

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    build_slide_1_cover(prs, slots)
    build_slide_2_welfare_overview(prs)
    build_slide_3_shanghai_skyline(prs, slots)
    build_slide_4_shanghai_culture(prs, slots)
    build_slide_5_hangzhou(prs, slots)
    build_slide_6_seafood_and_reflection(prs, slots)

    out = Path(output_path)
    prs.save(str(out))
    print(f"[完成] PPT 已生成：{out.resolve()}")
    return out


def main():
    parser = argparse.ArgumentParser(description="生成艾多美 5 周年员工旅游心得 PPT")
    parser.add_argument(
        "-p",
        "--photos",
        default=DEFAULT_MAC_PHOTO_DIR,
        help=f"照片文件夹路径（默认：{DEFAULT_MAC_PHOTO_DIR}）",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=DEFAULT_OUTPUT,
        help=f"输出 PPTX 文件路径（默认：{DEFAULT_OUTPUT}）",
    )
    args = parser.parse_args()
    generate_ppt(args.photos, args.output)


if __name__ == "__main__":
    main()
