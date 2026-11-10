#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
艾多美5周年员工旅游心得 PPT 生成脚本
------------------------------------
使用方法：
1. 默认扫描路径：/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片
   （也可将照片放入当前目录的 photos/ 文件夹，或通过 --photos 参数指定目录）
2. 运行：
   python3 generate_travel_ppt.py
   或者指定照片目录和输出文件名：
   python3 generate_travel_ppt.py --photos "/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片" --output "艾多美5周年员工旅游心得.pptx"

照片匹配说明：
- 如果照片文件名包含关键词（如“外滩”、“金茂”、“自助”、“豫园”、“城隍庙”、“南京路”、“西湖”、“雷峰塔”、“海鲜”、“龙虾”、“蟹”、“全家”、“合影”等），脚本会优先将对应照片智能分配到匹配的页面！
- 其余照片会按时间/文件名顺序自动填充到各页相册框中，并自动居中裁剪（不会拉伸变形）。
- 若未找到照片目录或照片数量不足，脚本会自动生成精美的占位插画图，方便后续在 PowerPoint / WPS 中右键“更改图片”。
"""

import argparse
import io
import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from PIL import Image, ImageDraw, ImageFont, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

# 默认照片目录（Mac 本地路径）
DEFAULT_PHOTO_DIR = "/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片"
FALLBACK_PHOTO_DIR = "photos"
DEFAULT_OUTPUT_PPTX = "艾多美5周年员工旅游心得_上海杭州之旅.pptx"

# 配色方案：艾多美蓝 + 江南暖金 + 雅致灰白
COLOR_ATOMY_BLUE = RGBColor(0x00, 0x95, 0xD9)      # 艾多美经典蓝
COLOR_DEEP_NAVY = RGBColor(0x0D, 0x2B, 0x45)       # 深海藏青（标题/深色背景）
COLOR_WARM_GOLD = RGBColor(0xD9, 0x9B, 0x26)       # 暖金点缀
COLOR_CORAL_RED = RGBColor(0xE0, 0x5A, 0x47)       # 美食/活力珊瑚红
COLOR_LAKE_TEAL = RGBColor(0x1E, 0x82, 0x7A)       # 西湖青绿
COLOR_BG_LIGHT = RGBColor(0xF6, 0xF8, 0xFB)        # 浅灰蓝背景
COLOR_CARD_WHITE = RGBColor(0xFF, 0xFF, 0xFF)      # 纯白卡片
COLOR_TEXT_DARK = RGBColor(0x1F, 0x2D, 0x3D)       # 正文深灰黑
COLOR_TEXT_MUTED = RGBColor(0x60, 0x6F, 0x7B)      # 辅助说明灰
COLOR_BORDER_LIGHT = RGBColor(0xE2, 0xE8, 0xF0)    # 浅边框

FONT_NAME = "Microsoft YaHei"  # 微软雅黑 / 苹方通用回退

SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff"}


def scan_photos(photo_dir: str) -> List[Path]:
    """扫描照片目录下的所有可用图片文件"""
    candidates = [Path(photo_dir), Path(FALLBACK_PHOTO_DIR)]
    for folder in candidates:
        if folder.exists() and folder.is_dir():
            files = [
                p for p in sorted(folder.rglob("*"))
                if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS and not p.name.startswith(".")
            ]
            if files:
                print(f"[✓] 已从目录加载 {len(files)} 张照片: {folder}")
                return files
    print(f"[!] 未在 '{photo_dir}' 或 '{FALLBACK_PHOTO_DIR}/' 找到照片，将使用精美主题占位图生成（可在PPT中直接右键替换照片）。")
    return []


def assign_photos_to_slots(photos: List[Path], slot_Specs: List[Tuple[str, List[str]]]) -> Dict[str, Optional[Path]]:
    """
    根据文件名关键词优先匹配槽位，未命中关键词的槽位按顺序分配剩余照片。
    如果照片总数少于槽位数，则循环使用照片。
    """
    assigned: Dict[str, Optional[Path]] = {slot_id: None for slot_id, _ in slot_Specs}
    if not photos:
        return assigned

    used_indices = set()

    # 第一轮：按文件名关键词匹配
    for slot_id, keywords in slot_Specs:
        for idx, p in enumerate(photos):
            if idx in used_indices:
                continue
            stem = p.stem.lower()
            if any(kw.lower() in stem for kw in keywords):
                assigned[slot_id] = p
                used_indices.add(idx)
                break

    # 第二轮：按顺序填充未匹配到的槽位
    remaining = [p for idx, p in enumerate(photos) if idx not in used_indices]
    rem_idx = 0
    for slot_id, _ in slot_Specs:
        if assigned[slot_id] is None:
            if rem_idx < len(remaining):
                assigned[slot_id] = remaining[rem_idx]
                rem_idx += 1
            else:
                # 照片不够时循环复用已有照片
                assigned[slot_id] = photos[rem_idx % len(photos)]
                rem_idx += 1

    return assigned


def _load_cjk_font(size: int) -> ImageFont.FreeTypeFont:
    """加载支持中文的字体用于生成占位图"""
    font_paths = [
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for fp in font_paths:
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                continue
    return ImageFont.load_default()


def prepare_cropped_image(
    photo_path: Optional[Path],
    target_w_inch: float,
    target_h_inch: float,
    placeholder_title: str,
    placeholder_subtitle: str,
    theme_rgb: Tuple[int, int, int] = (0, 149, 217),
) -> io.BytesIO:
    """
    将真实照片按目标宽高比居中裁剪（避免PPT中变形），并输出为内存字节流。
    如果 photo_path 为 None，则绘制一张高颜值的主题占位卡片图。
    """
    dpi = 220
    px_w = max(400, int(target_w_inch * dpi))
    px_h = max(300, int(target_h_inch * dpi))

    if photo_path and photo_path.exists():
        try:
            with Image.open(photo_path) as img:
                img = ImageOps.exif_transpose(img)
                img = img.convert("RGB")
                cropped = ImageOps.fit(img, (px_w, px_h), method=Image.Resampling.LANCZOS, centering=(0.5, 0.5))
                buf = io.BytesIO()
                cropped.save(buf, format="JPEG", quality=92)
                buf.seek(0)
                return buf
        except Exception as e:
            print(f"[!] 读取照片 {photo_path} 失败 ({e})，改用占位图。")

    # 生成精美的插画风占位图
    img = Image.new("RGB", (px_w, px_h), color=(240, 245, 250))
    draw = ImageDraw.Draw(img)

    r, g, b = theme_rgb
    # 绘制渐变感几何装饰背景
    for i in range(px_h):
        ratio = i / max(1, px_h - 1)
        cr = int(r * 0.14 + (248 - r * 0.14) * ratio)
        cg = int(g * 0.14 + (250 - g * 0.14) * ratio)
        cb = int(b * 0.14 + (253 - b * 0.14) * ratio)
        draw.line([(0, i), (px_w, i)], fill=(cr, cg, cb))

    # 装饰圆环与底部色带
    circle_r = int(min(px_w, px_h) * 0.28)
    cx, cy = px_w // 2, int(px_h * 0.42)
    draw.ellipse(
        [cx - circle_r, cy - circle_r, cx + circle_r, cy + circle_r],
        fill=(min(255, r + 35), min(255, g + 35), min(255, b + 35)),
        outline=(r, g, b),
        width=4,
    )

    # 相机图标简笔画
    cam_w, cam_h = int(circle_r * 0.7), int(circle_r * 0.5)
    draw.rounded_rectangle(
        [cx - cam_w // 2, cy - cam_h // 2, cx + cam_w // 2, cy + cam_h // 2],
        radius=10,
        fill=(255, 255, 255),
    )
    lens_r = int(cam_h * 0.3)
    draw.ellipse([cx - lens_r, cy - lens_r, cx + lens_r, cy + lens_r], fill=(r, g, b))

    # 底部色条
    bar_h = int(px_h * 0.26)
    draw.rectangle([0, px_h - bar_h, px_w, px_h], fill=(r, g, b))

    # 绘制文字说明
    font_title = _load_cjk_font(max(20, int(min(px_w, px_h) * 0.065)))
    font_sub = _load_cjk_font(max(14, int(min(px_w, px_h) * 0.042)))

    # 居中文字计算
    try:
        bbox1 = draw.textbbox((0, 0), placeholder_title, font=font_title)
        tw1 = bbox1[2] - bbox1[0]
        draw.text(((px_w - tw1) / 2, px_h - bar_h + int(bar_h * 0.18)), placeholder_title, fill=(255, 255, 255), font=font_title)

        bbox2 = draw.textbbox((0, 0), placeholder_subtitle, font=font_sub)
        tw2 = bbox2[2] - bbox2[0]
        draw.text(((px_w - tw2) / 2, px_h - bar_h + int(bar_h * 0.58)), placeholder_subtitle, fill=(235, 245, 255), font=font_sub)
    except Exception:
        draw.text((20, px_h - bar_h + 10), placeholder_title, fill=(255, 255, 255))

    # 外框虚线感内边
    pad = 12
    draw.rectangle([pad, pad, px_w - pad, px_h - pad], outline=(r, g, b), width=2)

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=90)
    buf.seek(0)
    return buf


# ================= PPT 绘制辅助工具函数 =================

def set_slide_background(slide, color: RGBColor):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: RGBColor,
    border_color: Optional[RGBColor] = None,
    border_pt: float = 1.0,
    rounded: bool = False,
):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(
        shape_type, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(border_pt)
    else:
        shape.line.fill.background()
    if rounded and hasattr(shape, "adjustments") and len(shape.adjustments) > 0:
        shape.adjustments[0] = 0.04
    return shape


def add_text_box(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    font_size: int = 16,
    color: RGBColor = COLOR_TEXT_DARK,
    bold: bool = False,
    align=PP_ALIGN.LEFT,
    line_spacing: float = 1.3,
):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)

    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(4)
        for run in p.runs:
            run.font.name = FONT_NAME
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.color.rgb = color
    return tb


def add_header(slide, page_num: str, title: str, subtitle: str, badge_text: str = "艾多美 · 五周年员工福利之旅"):
    """统一的顶部页眉设计"""
    # 顶部细装饰条
    add_rect(slide, 0, 0, 13.333, 0.10, COLOR_ATOMY_BLUE)

    # 页码徽章
    badge = add_rect(slide, 0.6, 0.35, 0.65, 0.55, COLOR_ATOMY_BLUE, rounded=True)
    tf = badge.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = page_num
    p.alignment = PP_ALIGN.CENTER
    for r in p.runs:
        r.font.name = FONT_NAME
        r.font.size = Pt(18)
        r.font.bold = True
        r.font.color.rgb = COLOR_CARD_WHITE

    # 标题与副标题
    add_text_box(slide, 1.4, 0.28, 8.2, 0.45, title, font_size=24, color=COLOR_DEEP_NAVY, bold=True)
    add_text_box(slide, 1.4, 0.68, 8.5, 0.32, subtitle, font_size=13, color=COLOR_TEXT_MUTED)

    # 右上角企业胶囊标签
    pill = add_rect(slide, 9.8, 0.38, 2.93, 0.42, COLOR_CARD_WHITE, border_color=COLOR_ATOMY_BLUE, border_pt=1.2, rounded=True)
    tf2 = pill.text_frame
    tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
    p2 = tf2.paragraphs[0]
    p2.text = badge_text
    p2.alignment = PP_ALIGN.CENTER
    for r in p2.runs:
        r.font.name = FONT_NAME
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = COLOR_ATOMY_BLUE

    # 分隔细线
    add_rect(slide, 0.6, 1.05, 12.133, 0.02, COLOR_BORDER_LIGHT)


def add_photo_frame(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    photo_path: Optional[Path],
    caption: str,
    placeholder_sub: str = "可在PPT中右键「更改图片」替换",
    theme_rgb: Tuple[int, int, int] = (0, 149, 217),
):
    """添加带白色相框衬底和底部图注的照片卡片"""
    # 白色相框卡片背景
    add_rect(
        slide,
        left,
        top,
        width,
        height,
        COLOR_CARD_WHITE,
        border_color=COLOR_BORDER_LIGHT,
        border_pt=1.2,
        rounded=True,
    )

    pad = 0.10
    caption_h = 0.40
    img_w = width - pad * 2
    img_h = height - pad * 2 - caption_h

    img_stream = prepare_cropped_image(
        photo_path,
        img_w,
        img_h,
        placeholder_title=caption,
        placeholder_subtitle=placeholder_sub,
        theme_rgb=theme_rgb,
    )
    slide.shapes.add_picture(
        img_stream,
        Inches(left + pad),
        Inches(top + pad),
        width=Inches(img_w),
        height=Inches(img_h),
    )

    # 图注文字
    add_text_box(
        slide,
        left + pad,
        top + height - caption_h - 0.03,
        img_w,
        caption_h,
        f"▲ {caption}",
        font_size=11,
        color=COLOR_TEXT_MUTED,
        bold=True,
        align=PP_ALIGN.CENTER,
    )


def add_bullet_card(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    tag_text: str,
    tag_color: RGBColor,
    card_title: str,
    bullets: List[str],
):
    """添加左侧带色标的结构化心得卡片"""
    add_rect(slide, left, top, width, height, COLOR_CARD_WHITE, border_color=COLOR_BORDER_LIGHT, border_pt=1.0, rounded=True)
    # 左侧竖向色条
    add_rect(slide, left, top, 0.10, height, tag_color)

    # 标签小胶囊
    tag = add_rect(slide, left + 0.25, top + 0.18, 1.55, 0.30, tag_color, rounded=True)
    tf_tag = tag.text_frame
    tf_tag.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = tag_text
    p_tag.alignment = PP_ALIGN.CENTER
    for r in p_tag.runs:
        r.font.name = FONT_NAME
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = COLOR_CARD_WHITE

    # 卡片标题
    add_text_box(slide, left + 1.90, top + 0.14, width - 2.05, 0.36, card_title, font_size=16, color=COLOR_DEEP_NAVY, bold=True)

    # 正文列表
    body_text = "\n".join(f"•  {b}" for b in bullets)
    add_text_box(
        slide,
        left + 0.25,
        top + 0.58,
        width - 0.40,
        height - 0.68,
        body_text,
        font_size=13,
        color=COLOR_TEXT_DARK,
        line_spacing=1.35,
    )


# ================= 各页面构建逻辑 =================

def build_presentation(photo_dir: str, output_path: str):
    prs = Presentation()
    # 16:9 宽屏 (13.333 x 7.5 英寸)
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    photos = scan_photos(photo_dir)

    # 定义全部 13 个照片槽位与智能匹配关键词
    slot_specs = [
        ("cover_main", ["全家", "合影", "家庭", "父母", "外滩", "西湖", "cover", "1"]),
        ("s2_family_1", ["父母", "爸妈", "老人", "家人", "合影", "出发", "高铁"]),
        ("s2_family_2", ["老婆", "孩子", "宝宝", "一家", "家庭", "笑容"]),
        ("s3_bund", ["外滩", "东方明珠", "陆家嘴", "江景", "黄浦江", "夜景"]),
        ("s3_jinmao_view", ["金茂", "56", "俯瞰", "高空", "观景", "大厦"]),
        ("s3_jinmao_buffet", ["自助", "金茂美食", "甜品", "牛排", "三文鱼", "餐厅"]),
        ("s4_yuyuan", ["豫园", "九曲桥", "亭", "园林", "古建"]),
        ("s4_chenghuang", ["城隍庙", "小吃", "生煎", "南翔", "小笼"]),
        ("s4_nanjing", ["南京路", "步行街", "和平饭店", "街景", "繁华"]),
        ("s5_westlake", ["西湖", "游船", "断桥", "三潭印月", "苏堤", "湖水"]),
        ("s5_leifeng", ["雷峰塔", "塔", "登高", "全景", "西湖全景"]),
        ("s6_lobster", ["龙虾", "海鲜", "大餐", "澳龙", "小青龙", "虾"]),
        ("s6_crab", ["面包蟹", "螃蟹", "大闸蟹", "蟹", "海鲜桌", "聚餐"]),
        ("s7_closing", ["合影", "全家福", "结束", "返程", "幸福", "感恩"]),
    ]
    slot_map = assign_photos_to_slots(photos, slot_specs)

    # -------------------------------------------------------------------------
    # 第 1 页：封面 —— 感恩五载 · 携爱同行（上海 & 杭州全家幸福之旅）
    # -------------------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1, COLOR_DEEP_NAVY)

    # 左侧品牌色块与金色装饰条
    add_rect(s1, 0, 0, 0.28, 7.5, COLOR_ATOMY_BLUE)
    add_rect(s1, 0.28, 0, 0.08, 7.5, COLOR_WARM_GOLD)

    # 左上角里程碑标签
    pill = add_rect(s1, 0.85, 0.85, 4.2, 0.48, COLOR_ATOMY_BLUE, rounded=True)
    tf = pill.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = "ATOMY · 艾多美 5周年员工福利旅游分享"
    p.alignment = PP_ALIGN.CENTER
    for r in p.runs:
        r.font.name = FONT_NAME
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = COLOR_CARD_WHITE

    # 主标题 & 副标题
    add_text_box(
        s1, 0.85, 1.65, 6.6, 1.8,
        "感恩五载 · 携爱同行\n上海 & 杭州全家幸福之旅",
        font_size=36, color=COLOR_CARD_WHITE, bold=True, line_spacing=1.25
    )
    add_text_box(
        s1, 0.85, 3.45, 6.4, 0.9,
        "公司承担费用 · 满 3 / 5 / 7 / 10 年专属员工旅游福利\n带着父母、老婆与孩子，共赴一场海派繁华与江南诗意之约",
        font_size=15, color=RGBColor(0xCB, 0xD5, 0xE1), line_spacing=1.45
    )

    # 行程亮点概览卡片（左下）
    info_cards = [
        ("满 5 周年", "艾多美暖心福利\n公司全额承担旅游费用", COLOR_WARM_GOLD),
        ("全家总动员", "父母 + 老婆 + 孩子\n三代同堂的温馨陪伴", COLOR_ATOMY_BLUE),
        ("双城经典游", "上海外滩/金茂56层/豫园\n杭州西湖/雷峰塔/海鲜大餐", COLOR_LAKE_TEAL),
    ]
    for idx, (head, desc, accent) in enumerate(info_cards):
        cx = 0.85 + idx * 2.15
        cy = 4.75
        add_rect(s1, cx, cy, 2.0, 1.85, RGBColor(0x16, 0x3A, 0x59), border_color=accent, border_pt=1.5, rounded=True)
        add_rect(s1, cx, cy, 2.0, 0.08, accent)
        add_text_box(s1, cx + 0.12, cy + 0.22, 1.76, 0.4, head, font_size=16, color=accent, bold=True)
        add_text_box(s1, cx + 0.12, cy + 0.72, 1.76, 0.95, desc, font_size=12, color=COLOR_CARD_WHITE, line_spacing=1.35)

    # 右侧大幅全家福相框
    add_photo_frame(
        s1, 7.45, 0.80, 5.25, 5.85,
        slot_map["cover_main"],
        "5周年全家幸福出游合影（上海 · 杭州）",
        theme_rgb=(0, 149, 217),
    )

    # -------------------------------------------------------------------------
    # 第 2 页：五年里程碑与福利初心 —— 一人努力，全家共享
    # -------------------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2, COLOR_BG_LIGHT)
    add_header(
        s2, "01",
        "五年同行 · 感恩公司的暖心旅游福利",
        "艾多美 3 / 5 / 7 / 10 周年员工旅游制度 —— 让工作的收获转化为家人的笑容"
    )

    # 顶部3-5-7-10周年时间轴阶梯条
    add_rect(s2, 0.6, 1.25, 12.133, 1.25, COLOR_CARD_WHITE, border_color=COLOR_BORDER_LIGHT, border_pt=1.0, rounded=True)
    add_text_box(s2, 0.85, 1.38, 3.2, 0.4, "艾多美员工旅游福利阶梯", font_size=15, color=COLOR_DEEP_NAVY, bold=True)
    add_text_box(s2, 0.85, 1.78, 3.2, 0.55, "每到一个年资里程碑，公司全额承担旅游费用，见证成长与蜕变", font_size=11, color=COLOR_TEXT_MUTED)

    milestones = [
        ("满 3 年", "初绽芳华 · 开启福利之旅", False),
        ("满 5 年 (本次)", "中流砥柱 · 携父母妻儿游沪杭", True),
        ("满 7 年", "笃行致远 · 期待下一程精彩", False),
        ("满 10 年", "十年功勋 · 荣耀梦想远航", False),
    ]
    for i, (m_title, m_sub, is_current) in enumerate(milestones):
        mx = 4.25 + i * 2.05
        my = 1.40
        bg_c = COLOR_ATOMY_BLUE if is_current else RGBColor(0xEB, 0xF3, 0xF9)
        txt_c = COLOR_CARD_WHITE if is_current else COLOR_DEEP_NAVY
        sub_c = RGBColor(0xE6, 0xF7, 0xFF) if is_current else COLOR_TEXT_MUTED
        bdr_c = COLOR_WARM_GOLD if is_current else COLOR_BORDER_LIGHT
        add_rect(s2, mx, my, 1.90, 0.95, bg_c, border_color=bdr_c, border_pt=2.0 if is_current else 1.0, rounded=True)
        add_text_box(s2, mx + 0.08, my + 0.12, 1.74, 0.35, m_title, font_size=13, color=txt_c, bold=True, align=PP_ALIGN.CENTER)
        add_text_box(s2, mx + 0.08, my + 0.48, 1.74, 0.40, m_sub, font_size=10, color=sub_c, align=PP_ALIGN.CENTER)

    # 左侧两张心得卡片
    add_bullet_card(
        s2, 0.6, 2.72, 6.0, 2.05,
        "感恩福利", COLOR_ATOMY_BLUE,
        "公司买单，让爱没有后顾之忧",
        [
            "入职艾多美满5年，迎来了期盼已久的五周年专属员工旅游福利。",
            "旅游费用由公司承担，真正体现了艾多美「珍爱灵魂、以人为本」的企业文化。",
            "不仅是对过去5年辛勤付出的认可，更是给整个家庭的一份厚礼。"
        ]
    )

    add_bullet_card(
        s2, 0.6, 4.95, 6.0, 2.05,
        "全家出动", COLOR_WARM_GOLD,
        "带父母、老婆、孩子共赴沪杭双城",
        [
            "平时工作忙碌，难得有机会带父母、老婆和孩子三代同堂一起远行。",
            "精选「上海 + 杭州」路线：既有魔都的摩登繁华，又有西湖的诗画江南，老少皆宜。",
            "看到父母欣慰的笑容、老婆放松的状态和孩子兴奋的眼神，倍感作为艾多美人的自豪！"
        ]
    )

    # 右侧两张家庭出游照片
    add_photo_frame(
        s2, 6.85, 2.72, 2.85, 4.28,
        slot_map["s2_family_1"],
        "陪伴父母踏上5周年福利旅程",
        theme_rgb=(217, 155, 38),
    )
    add_photo_frame(
        s2, 9.88, 2.72, 2.85, 4.28,
        slot_map["s2_family_2"],
        "老婆与孩子的灿烂笑容",
        theme_rgb=(0, 149, 217),
    )

    # -------------------------------------------------------------------------
    # 第 3 页：上海站（上）—— 外滩风云与金茂大厦56层云端观景自助
    # -------------------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3, COLOR_BG_LIGHT)
    add_header(
        s3, "02",
        "第一站 · 魔都上海：外滩风光与金茂56层云端自助",
        "漫步万国建筑博览群，登顶金茂大厦56层享用 300元/位 高空观景自助大餐"
    )

    # 左侧阐述卡片
    add_bullet_card(
        s3, 0.6, 1.30, 5.35, 2.65,
        "上海地标", COLOR_ATOMY_BLUE,
        "外滩漫步 · 感受十里洋场的璀璨气魄",
        [
            "全家一起漫步黄浦江畔的外滩，吹着江风，对岸陆家嘴摩天大楼群尽收眼底。",
            "父母看着电视里常出现的东方明珠与外滩钟楼就在眼前，连连赞叹上海的繁荣与气派。",
            "一家人在外滩留下珍贵合影，定格五周年最美的海派记忆。"
        ]
    )

    add_bullet_card(
        s3, 0.6, 4.15, 5.35, 2.85,
        "云端盛宴", COLOR_CORAL_RED,
        "金茂大厦 56 层 · 300元/位 观景自助餐",
        [
            "特意带全家打卡金茂大厦56层高空餐厅，享受 300元/位 的豪华观景自助餐。",
            "仰望震撼的金色中庭旋转回廊，俯瞰浦江两岸城市天际线，视觉与味觉双重震撼。",
            "琳琅满目的环球美食、精致甜品与鲜美料理，让父母和孩子吃得津津有味，仪式感拉满！"
        ]
    )

    # 右侧 3 张照片拼图（1大 + 2小）
    add_photo_frame(
        s3, 6.20, 1.30, 6.53, 3.05,
        slot_map["s3_bund"],
        "上海外滩 · 全家漫步黄浦江畔打卡留念",
        theme_rgb=(13, 43, 69),
    )
    add_photo_frame(
        s3, 6.20, 4.50, 3.18, 2.50,
        slot_map["s3_jinmao_view"],
        "金茂大厦56层 · 震撼高空观景视野",
        theme_rgb=(217, 155, 38),
    )
    add_photo_frame(
        s3, 9.55, 4.50, 3.18, 2.50,
        slot_map["s3_jinmao_buffet"],
        "300元/位 云端观景自助餐美食时刻",
        theme_rgb=(224, 90, 71),
    )

    # -------------------------------------------------------------------------
    # 第 4 页：上海站（下）—— 豫园雅韵、城隍庙烟火与南京路繁华
    # -------------------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4, COLOR_BG_LIGHT)
    add_header(
        s4, "03",
        "海派古今 · 豫园、城隍庙与南京路步行街",
        "在江南古典园林与百年商业街之间，品味老上海的传统韵味与现代摩登"
    )

    # 上方三列照片画廊
    spots_s4 = [
        ("s4_yuyuan", "豫园 · 江南古典园林亭台楼阁", (30, 130, 122), "一步一景，父母最爱的古典江南神韵"),
        ("s4_chenghuang", "城隍庙 · 飞檐翘角与地道海派小吃", (217, 155, 38), "九曲桥畔热闹非凡，尝遍特色点心"),
        ("s4_nanjing", "南京路步行街 · 中华商业第一街", (0, 149, 217), "霓虹闪烁、百年老字号与现代潮流交汇"),
    ]
    for i, (slot_key, cap, rgb, sub_desc) in enumerate(spots_s4):
        lx = 0.6 + i * 4.16
        add_photo_frame(
            s4, lx, 1.28, 3.82, 3.55,
            slot_map[slot_key],
            cap,
            theme_rgb=rgb,
        )

    # 下方综合游览心得卡片
    add_bullet_card(
        s4, 0.6, 5.02, 12.133, 2.00,
        "老上海风情", COLOR_WARM_GOLD,
        "古色古香与十里繁华交织的温馨慢时光",
        [
            "【豫园 & 城隍庙】带父母走过九曲桥，欣赏豫园精巧绝伦的江南亭台与假山锦鲤；在城隍庙品尝正宗南翔小笼包与传统糕点，长辈们对古建筑与老字号赞不绝口。",
            "【南京路步行街】伴着铛铛车的铃声漫步南京路，老婆和孩子在潮流店铺与老字号之间开心挑选伴手礼，一家人边走边逛，其乐融融。"
        ]
    )

    # -------------------------------------------------------------------------
    # 第 5 页：杭州站 —— 西湖泛舟与雷峰塔登高望远
    # -------------------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5, COLOR_BG_LIGHT)
    add_header(
        s5, "04",
        "第二站 · 诗画杭州：漫步西湖山水与登临雷峰塔",
        "「欲把西湖比西子，淡妆浓抹总相宜」—— 带家人沉浸在人间天堂的湖光山色中"
    )

    # 左侧两张大幅风景/合影照片
    add_photo_frame(
        s5, 0.6, 1.30, 3.65, 5.70,
        slot_map["s5_westlake"],
        "杭州西湖 · 碧波荡漾与苏堤春晓美景",
        theme_rgb=(30, 130, 122),
    )
    add_photo_frame(
        s5, 4.45, 1.30, 3.65, 5.70,
        slot_map["s5_leifeng"],
        "雷峰塔 · 登高远眺西湖全景与千年传奇",
        theme_rgb=(217, 155, 38),
    )

    # 右侧两张心得卡片
    add_bullet_card(
        s5, 8.30, 1.30, 4.43, 2.75,
        "西湖漫步", COLOR_LAKE_TEAL,
        "人在画中游 · 惬意江南慢生活",
        [
            "告别魔都的快节奏，来到烟雨温润的杭州西湖。",
            "垂柳依依、碧波荡漾，陪伴父母和妻儿沿着湖畔漫步，呼吸清新空气。",
            "孩子对西湖的游船与诗词故事充满好奇，旅途亦是最好的亲子课堂。"
        ]
    )

    add_bullet_card(
        s5, 8.30, 4.25, 4.43, 2.75,
        "雷峰夕照", COLOR_WARM_GOLD,
        "登临雷峰塔 · 饱览西湖壮丽全景",
        [
            "全家人乘坐扶梯轻松登临雷峰塔（对老人和孩子非常友好）。",
            "站在塔顶凭栏远眺，整个西湖的湖光山色、三潭印月与杭州城景尽收眼底。",
            "聆听千年白蛇传说的文化底蕴，感受历史古迹与自然山水的完美交融。"
        ]
    )

    # -------------------------------------------------------------------------
    # 第 6 页：舌尖上的幸福 —— 鲜香盛宴：龙虾与面包蟹海鲜大餐
    # -------------------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6, COLOR_BG_LIGHT)
    add_header(
        s6, "05",
        "舌尖上的旅途 · 龙虾与面包蟹海鲜大餐",
        "唯有爱与美食不可辜负 —— 用一桌丰盛鲜美的海鲜盛宴犒劳最亲爱的家人"
    )

    # 左侧两张高清美食照片
    add_photo_frame(
        s6, 0.6, 1.30, 3.85, 5.70,
        slot_map["s6_lobster"],
        "鲜活大龙虾 · 肉质Q弹鲜甜饱满",
        theme_rgb=(224, 90, 71),
    )
    add_photo_frame(
        s6, 4.65, 1.30, 3.85, 5.70,
        slot_map["s6_crab"],
        "膏满黄肥面包蟹 · 全家共享海鲜盛宴",
        theme_rgb=(217, 155, 38),
    )

    # 右侧美食心得与菜单亮点卡片
    add_bullet_card(
        s6, 8.70, 1.30, 4.03, 3.45,
        "饕餮海鲜", COLOR_CORAL_RED,
        "龙虾 + 面包蟹 · 满桌鲜香的团圆味",
        [
            "旅途中最热闹的时刻，莫过于全家人围坐一桌享用海鲜大餐！",
            "个头硕大的大龙虾肉质紧实弹牙，饱满的面包蟹蟹膏丰腴、蟹肉鲜甜。",
            "从金茂大厦56层300元/位的云端自助，到地道江南风味与豪华海鲜大餐，这一路不仅饱了眼福，更让父母、老婆和孩子大饱口福！"
        ]
    )

    # 右下美食幸福感指数小卡片
    add_rect(s6, 8.70, 4.95, 4.03, 2.05, COLOR_DEEP_NAVY, rounded=True)
    add_rect(s6, 8.70, 4.95, 4.03, 0.08, COLOR_WARM_GOLD)
    add_text_box(s6, 8.95, 5.15, 3.55, 0.38, "旅途美食幸福记忆清单", font_size=15, color=COLOR_WARM_GOLD, bold=True)
    food_summary = (
        "★  上海金茂大厦 56 层观景自助（300元/位）\n"
        "★  城隍庙 & 南京路地道海派传统小吃\n"
        "★  豪华海鲜大餐：大龙虾、膏满面包蟹等\n"
        "♥  最大的美味：全家人整整齐齐举杯欢笑的时刻"
    )
    add_text_box(s6, 8.95, 5.58, 3.60, 1.30, food_summary, font_size=12, color=COLOR_CARD_WHITE, line_spacing=1.38)

    # -------------------------------------------------------------------------
    # 第 7 页：旅游心得总结与展望 —— 感恩艾多美，奔赴7年与10年之约
    # -------------------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7, COLOR_BG_LIGHT)
    add_header(
        s7, "06",
        "旅游心得感悟 · 感恩艾多美，蓄力再出发",
        "五年深耕不负韶华，带着家人的支持与祝福，迈向 7 周年、10 周年的更高山峰"
    )

    # 左侧三段核心感悟卡片
    insights = [
        (
            "感恩公司", COLOR_ATOMY_BLUE,
            "1. 一份有温度的福利，连接企业与家庭",
            "艾多美 3、5、7、10 周年的员工旅游福利，不仅由公司承担费用减轻了员工负担，更让我们有机会把工作的荣光带回家。父母在旅途中骄傲地向亲友提起「这是孩子公司满5年奖励的旅游」，这份认可是最珍贵的礼物。"
        ),
        (
            "亲情陪伴", COLOR_LAKE_TEAL,
            "2. 最美的风景，是全家人在一起的笑容",
            "从上海外滩、金茂56层云端自助、豫园城隍庙南京路，到杭州西湖、雷峰塔与海鲜大餐，父母的康健欢笑、老婆的温柔相伴、孩子的开阔眼界，让这次五周年之旅成为全家心中温暖而滚烫的记忆。"
        ),
        (
            "未来展望", COLOR_WARM_GOLD,
            "3. 化感恩为动力，奔赴 7 周年与 10 周年之约",
            "短暂的休整是为了更有力地出发！回到工作岗位后，我将把这份满满的幸福感转化为工作热情，与艾多美并肩同行，共同期待满 7 年、满 10 年时带家人去看更广阔的世界！"
        ),
    ]

    for idx, (tag_t, tag_c, title_t, desc_t) in enumerate(insights):
        top_y = 1.28 + idx * 1.92
        add_bullet_card(
            s7, 0.6, top_y, 7.35, 1.76,
            tag_t, tag_c, title_t, [desc_t]
        )

    # 右侧照片 + 结语寄语框
    add_photo_frame(
        s7, 8.20, 1.28, 4.53, 3.95,
        slot_map["s7_closing"],
        "感恩艾多美 · 五周年全家圆满之旅",
        theme_rgb=(0, 149, 217),
    )

    add_rect(s7, 8.20, 5.42, 4.53, 1.58, COLOR_DEEP_NAVY, rounded=True)
    add_rect(s7, 8.20, 5.42, 0.10, 1.58, COLOR_ATOMY_BLUE)
    add_text_box(
        s7, 8.45, 5.56, 4.15, 0.40,
        "五年同心 · 感恩有你 · 未来可期",
        font_size=16, color=COLOR_WARM_GOLD, bold=True, align=PP_ALIGN.CENTER
    )
    add_text_box(
        s7, 8.45, 6.02, 4.15, 0.85,
        "感谢艾多美公司提供如此暖心的5周年旅游福利！\n祝公司蒸蒸日上，我们7周年、10周年再会！",
        font_size=12, color=COLOR_CARD_WHITE, align=PP_ALIGN.CENTER, line_spacing=1.4
    )

    prs.save(output_path)
    print(f"[✓] PPT 已成功生成：{os.path.abspath(output_path)} （共 {len(prs.slides)} 页）")


def main():
    parser = argparse.ArgumentParser(description="生成艾多美5周年员工旅游心得PPT")
    parser.add_argument(
        "--photos",
        default=DEFAULT_PHOTO_DIR,
        help=f"照片文件夹路径（默认：{DEFAULT_PHOTO_DIR}）",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT_PPTX,
        help=f"输出的PPT文件名（默认：{DEFAULT_OUTPUT_PPTX}）",
    )
    args = parser.parse_args()
    build_presentation(args.photos, args.output)


if __name__ == "__main__":
    main()
