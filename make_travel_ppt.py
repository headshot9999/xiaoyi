#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
艾多美5周年员工旅游心得 PPT 生成器
==================================

将上海、杭州全家游的心得与照片自动排版成 16:9 宽屏 PPT（共 6 页）：
  P1 封面：五年同行，感恩有家 —— 艾多美5周年员工旅游心得
  P2 缘起与同行：公司3/5/7/10年旅游福利，携父母、妻子、孩子三代同游
  P3 上海站·云端与外滩：漫步外滩、南京路，金茂大厦56层观景自助（300元/位）
  P4 上海站·江南烟火：游览豫园、城隍庙，感受老上海历史底蕴与市井温情
  P5 杭州站·湖光山色与美食：西湖、雷峰塔，共享龙虾、面包蟹等海鲜大餐
  P6 旅游心得与感恩寄语：感恩艾多美平台，把奋斗化作给家人的幸福时光

用法：
  # 默认会自动读取 /Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片
  python3 make_travel_ppt.py

  # 也可手动指定照片目录与输出文件名
  python3 make_travel_ppt.py -i "/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片" -o "艾多美5周年员工旅游心得.pptx"
"""

import argparse
import io
import os
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except Exception:
    pass

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# 默认照片目录（支持您的本地 Mac 路径及当前目录下的 5周年旅游照片 文件夹）
DEFAULT_PHOTO_DIRS = [
    "/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片",
    "5周年旅游照片",
    "photos",
]

DEFAULT_OUTPUT = "艾多美5周年员工旅游心得.pptx"

# 字体与配色（艾多美蓝 + 温暖金橙点缀 + 干净米白底色）
FONT_SANS = "PingFang SC"
FONT_BOLD = "PingFang SC Semibold"

BG_MAIN = "F8FAFC"       # 柔和浅灰蓝底色
BG_WARM = "FAF8F5"       # 温暖米白底色
CARD_WHITE = "FFFFFF"    # 纯白卡片
ATOMY_BLUE = "0066A4"    # 艾多美品牌蓝
ATOMY_LIGHT = "EBF5FB"   # 浅艾多美蓝底色
ATOMY_DARK = "0B2545"    # 深海蓝标题色
ACCENT_GOLD = "D97706"   # 暖金点缀色
ACCENT_SOFT = "FEF3C7"   # 暖金浅底色
TEXT_INK = "1E293B"      # 主正文深灰
TEXT_BODY = "334155"     # 正文色
TEXT_MUTED = "64748B"    # 辅助说明文字
BORDER_LINE = "E2E8F0"   # 细边框色

SLIDE_W = 13.333333
SLIDE_H = 7.5

# 每个照片槽位的主题关键词（用于按文件名自动匹配对应景点照片）
SLOT_SPECS = [
    {
        "id": "cover",
        "title": "全家福 / 旅途合影",
        "subtitle": "上海 · 杭州 五周年全家游主图",
        "keywords": ["全家", "合影", "封面", "家人", "父母", "一家", "合照", "cover", "family"],
        "theme_color": (11, 37, 69),
    },
    {
        "id": "family",
        "title": "三代同游 · 幸福同行",
        "subtitle": "父母 · 妻子 · 孩子 旅途留念",
        "keywords": ["父母", "爸妈", "老爸", "老妈", "老婆", "妻子", "孩子", "宝宝", "出发", "高铁", "家人"],
        "theme_color": (0, 102, 164),
    },
    {
        "id": "bund",
        "title": "上海外滩 · 南京路步行街",
        "subtitle": "浦江两岸风光与十里洋场繁华",
        "keywords": ["外滩", "南京路", "东方明珠", "黄浦江", "上海", "bund", "步行街"],
        "theme_color": (20, 66, 114),
    },
    {
        "id": "jinmao",
        "title": "金茂大厦 56层 · 观景自助餐",
        "subtitle": "300元/位 · 云端俯瞰陆家嘴与全家人共享美食",
        "keywords": ["金茂", "56", "自助", "云端", "天际", "餐厅", "jinmao", "陆家嘴"],
        "theme_color": (161, 98, 7),
    },
    {
        "id": "yuyuan",
        "title": "上海豫园 · 江南古典园林",
        "subtitle": "亭台楼阁 · 九曲桥畔家人漫步",
        "keywords": ["豫园", "九曲桥", "园林", "yuyuan", "亭"],
        "theme_color": (15, 118, 110),
    },
    {
        "id": "chenghuang",
        "title": "老城隍庙 · 传统市井烟火",
        "subtitle": "古色古香街区与地道上海风味",
        "keywords": ["城隍庙", "小吃", "南翔", "小笼", "古街", "庙"],
        "theme_color": (180, 83, 9),
    },
    {
        "id": "westlake",
        "title": "杭州西湖 · 雷峰塔",
        "subtitle": "湖光山色 · 雷峰夕照与西子湖畔合影",
        "keywords": ["西湖", "雷峰塔", "杭州", "断桥", "三潭印月", "游船", "westlake", "塔"],
        "theme_color": (21, 94, 117),
    },
    {
        "id": "seafood",
        "title": "丰盛海鲜大餐 · 龙虾与面包蟹",
        "subtitle": "鲜美大龙虾、膏满面包蟹，全家人大快朵颐",
        "keywords": ["海鲜", "龙虾", "面包蟹", "螃蟹", "大餐", "美食", "蟹", "虾", "吃"],
        "theme_color": (185, 28, 28),
    },
    {
        "id": "closing",
        "title": "五年相伴 · 感恩艾多美",
        "subtitle": "温馨旅途瞬间 · 期待7周年、10周年再出发",
        "keywords": ["温馨", "背影", "微笑", "风景", "开心", "纪念", "总结"],
        "theme_color": (0, 102, 164),
    },
]


def rgb(hex_str):
    return RGBColor.from_string(hex_str)


def set_text_frame_margins(tf, valign="top", margin_pt=0):
    body = tf._txBody.find(qn("a:bodyPr"))
    body.set("anchor", {"top": "t", "middle": "ctr", "bottom": "b"}[valign])
    emu_val = str(int(Pt(margin_pt)))
    for key in ("lIns", "rIns", "tIns", "bIns"):
        body.set(key, emu_val)


def add_run_styled(p, text, size=14, color=TEXT_INK, bold=False, font=FONT_SANS):
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
    return run


def add_textbox(slide, x, y, w, h, paragraphs, valign="top", margin_pt=0):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    set_text_frame_margins(tf, valign=valign, margin_pt=margin_pt)
    align_map = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
    }
    for idx, para_info in enumerate(paragraphs):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align_map.get(para_info.get("align", "left"), PP_ALIGN.LEFT)
        if "line" in para_info:
            p.line_spacing = para_info["line"]
        if para_info.get("before"):
            p.space_before = Pt(para_info["before"])
        if para_info.get("after"):
            p.space_after = Pt(para_info["after"])

        if "runs" in para_info:
            for r in para_info["runs"]:
                add_run_styled(
                    p,
                    r["text"],
                    size=r.get("size", 14),
                    color=r.get("color", TEXT_INK),
                    bold=r.get("bold", False),
                    font=r.get("font", FONT_BOLD if r.get("bold") else FONT_SANS),
                )
        else:
            bold = para_info.get("bold", False)
            add_run_styled(
                p,
                para_info["text"],
                size=para_info.get("size", 14),
                color=para_info.get("color", TEXT_INK),
                bold=bold,
                font=para_info.get("font", FONT_BOLD if bold else FONT_SANS),
            )
    return shape


def add_rect(slide, x, y, w, h, fill_hex, line_hex=None, line_pt=1.0, rounded=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded and hasattr(shape, "adjustments") and len(shape.adjustments) > 0:
        shape.adjustments[0] = 0.03
    if fill_hex is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill_hex)
    if line_hex is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(line_hex)
        shape.line.width = Pt(line_pt)
    return shape


def discover_photos(input_path=None):
    """查找并返回照片文件列表（按文件名与修改时间排序）。"""
    exts = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".heic", ".heif", ".tif", ".tiff"}
    search_dirs = []
    if input_path:
        search_dirs.append(input_path)
    for d in DEFAULT_PHOTO_DIRS:
        if d not in search_dirs:
            search_dirs.append(d)

    chosen_dir = None
    files = []
    for d in search_dirs:
        p = Path(d).expanduser()
        if p.is_file() and p.suffix.lower() in exts:
            return [p], p.parent
        if p.is_dir():
            found = sorted(
                [f for f in p.rglob("*") if f.is_file() and f.suffix.lower() in exts and not f.name.startswith(".")],
                key=lambda x: x.name,
            )
            if found:
                chosen_dir = p
                files = found
                break
            elif chosen_dir is None:
                chosen_dir = p

    return files, chosen_dir


def match_photos_to_slots(photo_files):
    """
    将用户照片智能分配到 9 个展示位置：
    1. 优先根据文件名中的关键词（如"外滩"、"金茂"、"豫园"、"西湖"、"龙虾"等）精准匹配；
    2. 其余照片按顺序填入剩余位置；
    3. 若照片少于 9 张，则循环使用已有照片，确保每一页都有真实照片展示。
    """
    if not photo_files:
        return {spec["id"]: None for spec in SLOT_SPECS}

    assigned = {}
    used_files = set()

    # 第一轮：文件名关键词匹配
    for spec in SLOT_SPECS:
        slot_id = spec["id"]
        for f in photo_files:
            if f in used_files:
                continue
            fname_lower = f.stem.lower()
            if any(kw.lower() in fname_lower for kw in spec["keywords"]):
                assigned[slot_id] = f
                used_files.add(f)
                break

    # 第二轮：剩余未使用的照片按顺序填入空位
    remaining_files = [f for f in photo_files if f not in used_files]
    rem_idx = 0
    for spec in SLOT_SPECS:
        slot_id = spec["id"]
        if slot_id not in assigned:
            if rem_idx < len(remaining_files):
                assigned[slot_id] = remaining_files[rem_idx]
                rem_idx += 1
            else:
                # 照片总数不足 9 张时，循环复用照片
                assigned[slot_id] = photo_files[len(assigned) % len(photo_files)]

    return assigned


def _load_pil_font(size_px):
    """加载系统可用的中文字体用于生成精美占位示意图。"""
    candidates = [
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
        "C:/Windows/Fonts/msyh.ttc",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size_px)
            except Exception:
                pass
    return ImageFont.load_default()


def render_photo_or_card(photo_path, width_in, height_in, spec, dpi=220):
    """
    将真实照片按目标宽高居中裁剪（自动修正手机旋转方向），
    若当前环境尚未放入照片，则生成高质感主题示意卡片。
    """
    target_w = max(200, int(round(width_in * dpi)))
    target_h = max(150, int(round(height_in * dpi)))

    if photo_path is not None and Path(photo_path).is_file():
        try:
            img = Image.open(photo_path)
            img = ImageOps.exif_transpose(img)
            img = img.convert("RGB")
            cropped = ImageOps.fit(img, (target_w, target_h), method=Image.LANCZOS, centering=(0.5, 0.45))
            buf = io.BytesIO()
            cropped.save(buf, format="JPEG", quality=92, dpi=(dpi, dpi))
            buf.seek(0)
            return buf
        except Exception as exc:
            print(f"  [提示] 读取照片 {photo_path} 失败 ({exc})，将使用主题示意卡代替。", file=sys.stderr)

    # 生成视觉优雅的主题景点示意卡（当在云端环境运行、未连接本地 Mac 照片目录时使用）
    base_r, base_g, base_b = spec["theme_color"]
    img = Image.new("RGB", (target_w, target_h), (base_r, base_g, base_b))
    draw = ImageDraw.Draw(img)

    # 绘制柔和的几何装饰背景
    for i in range(target_h):
        ratio = i / max(1, target_h - 1)
        r = int(base_r * (1 - 0.28 * ratio) + 18 * ratio)
        g = int(base_g * (1 - 0.28 * ratio) + 30 * ratio)
        b = int(base_b * (1 - 0.28 * ratio) + 48 * ratio)
        draw.line([(0, i), (target_w, i)], fill=(r, g, b))

    # 装饰线框与内边距
    pad = max(18, int(min(target_w, target_h) * 0.05))
    draw.rectangle(
        [pad, pad, target_w - pad, target_h - pad],
        outline=(255, 255, 255),
        width=2,
    )

    # 绘制顶部相机图标标识与文字说明
    font_title = _load_pil_font(max(22, int(min(target_w, target_h) * 0.068)))
    font_sub = _load_pil_font(max(15, int(min(target_w, target_h) * 0.042)))
    font_hint = _load_pil_font(max(13, int(min(target_w, target_h) * 0.034)))

    cx, cy = target_w // 2, target_h // 2

    # 相机小图标轮廓
    cam_w, cam_h = int(min(target_w, target_h) * 0.16), int(min(target_w, target_h) * 0.11)
    cam_y = cy - int(target_h * 0.16)
    draw.rounded_rectangle(
        [cx - cam_w // 2, cam_y - cam_h // 2, cx + cam_w // 2, cam_y + cam_h // 2],
        radius=8,
        outline=(255, 255, 255),
        width=3,
    )
    lens_r = int(cam_h * 0.3)
    draw.ellipse(
        [cx - lens_r, cam_y - lens_r, cx + lens_r, cam_y + lens_r],
        outline=(255, 255, 255),
        width=3,
    )

    # 居中绘制标题与副标题
    title_text = spec["title"]
    sub_text = spec["subtitle"]
    hint_text = "运行 python3 make_travel_ppt.py 自动载入『5周年旅游照片』文件夹实拍图"

    def draw_centered(text, y_pos, font, fill):
        bbox = draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        draw.text((cx - tw // 2, y_pos - th // 2), text, font=font, fill=fill)

    draw_centered(title_text, cy + int(target_h * 0.02), font_title, (255, 255, 255))
    draw_centered(sub_text, cy + int(target_h * 0.13), font_sub, (226, 232, 240))
    draw_centered(hint_text, target_h - pad - int(target_h * 0.06), font_hint, (190, 210, 230))

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=92, dpi=(dpi, dpi))
    buf.seek(0)
    return buf


def insert_framed_photo(slide, x, y, w, h, photo_path, spec, caption=None):
    """在幻灯片指定位置插入带白色相框衬底和底部说明条的照片。"""
    # 相框白底卡片
    add_rect(slide, x - 0.08, y - 0.08, w + 0.16, h + (0.48 if caption else 0.16), CARD_WHITE, BORDER_LINE, 1.0)
    img_buf = render_photo_or_card(photo_path, w, h, spec)
    slide.shapes.add_picture(img_buf, Inches(x), Inches(y), width=Inches(w), height=Inches(h))
    if caption:
        add_textbox(
            slide,
            x,
            y + h + 0.06,
            w,
            0.32,
            [{"text": caption, "size": 11, "color": TEXT_MUTED, "align": "center"}],
            valign="middle",
        )


def add_header(slide, page_tag, title, subtitle, page_num_str):
    """统一添加内页顶部栏与页码。"""
    # 顶部艾多美蓝细饰条
    add_rect(slide, 0, 0, SLIDE_W, 0.08, ATOMY_BLUE)

    # 章节胶囊标签
    add_rect(slide, 0.65, 0.36, 1.95, 0.34, ATOMY_LIGHT, None, rounded=True)
    add_textbox(
        slide,
        0.65,
        0.36,
        1.95,
        0.34,
        [{"text": page_tag, "size": 11, "color": ATOMY_BLUE, "bold": True, "align": "center"}],
        valign="middle",
    )

    # 页码与品牌标识
    add_textbox(
        slide,
        SLIDE_W - 3.6,
        0.36,
        2.95,
        0.34,
        [
            {
                "text": f"ATOMY 5周年员工旅游心得  |  {page_num_str}",
                "size": 10.5,
                "color": TEXT_MUTED,
                "align": "right",
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
        0.55,
        [{"text": title, "size": 26, "color": ATOMY_DARK, "bold": True}],
    )
    add_textbox(
        slide,
        0.65,
        1.32,
        12.0,
        0.36,
        [{"text": subtitle, "size": 13.5, "color": TEXT_MUTED}],
    )
    # 标题下分割细线
    add_rect(slide, 0.65, 1.74, 12.03, 0.015, BORDER_LINE)


# ===========================================================================
# 第 1 页：封面页
# ===========================================================================
def build_slide_1_cover(prs, blank_layout, slot_photos, specs_by_id):
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_WARM)

    # 左侧深色品牌视觉条
    add_rect(slide, 0, 0, 0.22, SLIDE_H, ATOMY_BLUE)

    # 左半区：文字内容
    add_rect(slide, 0.75, 0.85, 2.8, 0.40, ATOMY_LIGHT, None, rounded=True)
    add_textbox(
        slide,
        0.75,
        0.85,
        2.8,
        0.40,
        [{"text": "艾多美 · 入职满 5 周年员工福利", "size": 12, "color": ATOMY_BLUE, "bold": True, "align": "center"}],
        valign="middle",
    )

    add_textbox(
        slide,
        0.75,
        1.50,
        5.8,
        1.65,
        [
            {"text": "五年同行，感恩有家", "size": 38, "color": ATOMY_DARK, "bold": True},
            {"text": "上海 · 杭州 全家幸福之旅心得分享", "size": 20, "color": ATOMY_BLUE, "bold": True, "before": 14},
        ],
    )

    add_rect(slide, 0.75, 3.25, 1.1, 0.06, ACCENT_GOLD)

    add_textbox(
        slide,
        0.75,
        3.52,
        5.55,
        1.45,
        [
            {
                "text": "在艾多美，每满 3、5、7、10 年的员工旅游福利，不仅是一份旅途经费的支持，更是公司对员工与家人最温情的关怀。借着入职满 5 周年的契机，我带着父母、妻子和孩子奔赴上海与杭州，度过了一段难忘的团圆时光。",
                "size": 13.5,
                "color": TEXT_BODY,
                "line": 1.45,
            }
        ],
    )

    # 底部三个行程概览信息卡
    cards = [
        ("同行家人", "父母 · 妻子 · 孩子", "三代同堂，共享天伦之乐"),
        ("上海站足迹", "外滩 · 南京路 · 豫园 · 城隍庙", "金茂大厦56层观景自助（300元/位）"),
        ("杭州站足迹", "西湖风光 · 雷峰塔登高", "龙虾、面包蟹等丰盛海鲜大餐"),
    ]
    for i, (label, main_txt, sub_txt) in enumerate(cards):
        cy = 5.15 + i * 0.64
        add_rect(slide, 0.75, cy, 5.55, 0.54, CARD_WHITE, BORDER_LINE, 1.0, rounded=True)
        add_rect(slide, 0.75, cy, 0.08, 0.54, ATOMY_BLUE if i != 1 else ACCENT_GOLD)
        add_textbox(
            slide,
            0.98,
            cy + 0.06,
            1.35,
            0.42,
            [{"text": label, "size": 11.5, "color": ATOMY_BLUE if i != 1 else ACCENT_GOLD, "bold": True}],
            valign="middle",
        )
        add_textbox(
            slide,
            2.30,
            cy + 0.06,
            3.90,
            0.42,
            [
                {
                    "runs": [
                        {"text": main_txt + "  ", "size": 12, "color": TEXT_INK, "bold": True},
                        {"text": f"|  {sub_txt}", "size": 11, "color": TEXT_MUTED},
                    ]
                }
            ],
            valign="middle",
        )

    # 右半区：大幅全家福主图
    insert_framed_photo(
        slide,
        6.72,
        0.78,
        5.92,
        5.55,
        slot_photos["cover"],
        specs_by_id["cover"],
        caption="▲ 艾多美 5 周年员工旅游 · 携父母、妻子与孩子同游沪杭留念",
    )


# ===========================================================================
# 第 2 页：福利制度与三代同行
# ===========================================================================
def build_slide_2_welfare(prs, blank_layout, slot_photos, specs_by_id):
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_MAIN)
    add_header(
        slide,
        "01 / 福利与初心",
        "五载耕耘迎福利，三代同堂共出发",
        "公司每满 3、5、7、10 年全额承担员工旅游费用 —— 让工作的成果化作陪伴家人的底气",
        "02 / 06",
    )

    # 顶部里程碑节点条（3年 / 5年[高亮] / 7年 / 10年）
    milestones = [
        ("满 3 周年", "初绽芳华 · 感恩相伴", False),
        ("满 5 周年（本次旅程）", "携父母妻儿 · 沪杭幸福游", True),
        ("满 7 周年", "笃定前行 · 持续成长", False),
        ("满 10 周年", "十年同心 · 共筑未来", False),
    ]
    for i, (m_title, m_desc, active) in enumerate(milestones):
        mx = 0.65 + i * 3.06
        my = 1.96
        mw = 2.88
        mh = 0.88
        bg = ATOMY_BLUE if active else CARD_WHITE
        border = ATOMY_BLUE if active else BORDER_LINE
        title_col = CARD_WHITE if active else ATOMY_DARK
        desc_col = "E0F2FE" if active else TEXT_MUTED

        add_rect(slide, mx, my, mw, mh, bg, border, 1.2, rounded=True)
        add_textbox(
            slide,
            mx + 0.18,
            my + 0.12,
            mw - 0.36,
            0.34,
            [{"text": m_title, "size": 14, "color": title_col, "bold": True}],
        )
        add_textbox(
            slide,
            mx + 0.18,
            my + 0.46,
            mw - 0.36,
            0.30,
            [{"text": m_desc, "size": 11.5, "color": desc_col}],
        )

    # 左下：三张人物与行程心得卡片
    left_cards = [
        (
            "感恩公司制度：实在的福利，温暖的归属感",
            "艾多美设立的 3、5、7、10 年员工旅游制度，费用由公司承担，真正体现了『珍视灵魂』与关爱员工家庭的企业文化。转眼入职满 5 年，这份福利让我能够毫无后顾之忧地为全家人安排一次高品质旅行。",
        ),
        (
            "带着父母、老婆、孩子：最圆满的旅行阵容",
            "平时工作繁忙，很少有机会集齐一家老小共同出游。这次特意带上辛劳半生的父母、默默支持我的老婆和活泼可爱的孩子，选定交通便利、风景人文俱佳的『上海 + 杭州』双城路线，兼顾老人舒适度与孩子的眼界。",
        ),
        (
            "行程亮点：吃好、看尽、慢节奏陪伴",
            "从上海外滩、南京路、金茂大厦 56 层 300元/位的云端自助，到豫园、城隍庙的江南韵味，再到杭州西湖、雷峰塔与丰盛的龙虾面包蟹海鲜大餐，每一站都充满了全家人的欢声笑语。",
        ),
    ]
    for i, (c_title, c_body) in enumerate(left_cards):
        cy = 3.05 + i * 1.32
        add_rect(slide, 0.65, cy, 6.45, 1.18, CARD_WHITE, BORDER_LINE, 1.0, rounded=True)
        add_rect(slide, 0.65, cy, 0.08, 1.18, ATOMY_BLUE if i != 1 else ACCENT_GOLD)
        add_textbox(
            slide,
            0.90,
            cy + 0.12,
            6.05,
            0.32,
            [{"text": c_title, "size": 14, "color": ATOMY_DARK, "bold": True}],
        )
        add_textbox(
            slide,
            0.90,
            cy + 0.44,
            6.05,
            0.66,
            [{"text": c_body, "size": 11.8, "color": TEXT_BODY, "line": 1.35}],
        )

    # 右下：三代同游照片
    insert_framed_photo(
        slide,
        7.38,
        3.08,
        5.22,
        3.40,
        slot_photos["family"],
        specs_by_id["family"],
        caption="▲ 父母、妻子与孩子同行 —— 全家人的笑容是最珍贵的礼物",
    )


# ===========================================================================
# 第 3 页：上海站（一）外滩、南京路与金茂大厦56层自助
# ===========================================================================
def build_slide_3_shanghai_modern(prs, blank_layout, slot_photos, specs_by_id):
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_MAIN)
    add_header(
        slide,
        "02 / 上海站 · 繁华与云端",
        "漫步外滩与南京路，登金茂56层享云端自助",
        "感受十里洋场的摩登气象，在 56 层高空与父母妻儿共进 300元/位 观景自助餐",
        "03 / 06",
    )

    # 左栏：外滩与南京路照片 + 心得卡
    insert_framed_photo(
        slide,
        0.73,
        1.98,
        5.75,
        3.15,
        slot_photos["bund"],
        specs_by_id["bund"],
        caption="▲ 漫步外滩与南京路步行街，感受上海滩的历史与现代交融",
    )

    add_rect(slide, 0.65, 5.68, 5.91, 1.38, CARD_WHITE, BORDER_LINE, 1.0, rounded=True)
    add_rect(slide, 0.65, 5.68, 0.08, 1.38, ATOMY_BLUE)
    add_textbox(
        slide,
        0.88,
        5.78,
        5.55,
        0.34,
        [{"text": "外滩万国建筑群 & 南京路步行街", "size": 14.5, "color": ATOMY_DARK, "bold": True}],
    )
    add_textbox(
        slide,
        0.88,
        6.14,
        5.55,
        0.84,
        [
            {
                "text": "陪着父母和妻儿走过繁华的南京路步行街，一路来到黄浦江畔的外滩。父母看着对岸高耸的陆家嘴天际线感叹时代发展的迅速，孩子则对江面穿梭的游轮兴奋不已，一家人在江风中留下珍贵合影。",
                "size": 11.8,
                "color": TEXT_BODY,
                "line": 1.38,
            }
        ],
    )

    # 右栏：金茂大厦56层观景自助照片 + 心得卡
    insert_framed_photo(
        slide,
        6.85,
        1.98,
        5.75,
        3.15,
        slot_photos["jinmao"],
        specs_by_id["jinmao"],
        caption="▲ 金茂大厦 56 层观景自助餐（300元/位），美食与高空美景兼得",
    )

    add_rect(slide, 6.77, 5.68, 5.91, 1.38, CARD_WHITE, BORDER_LINE, 1.0, rounded=True)
    add_rect(slide, 6.77, 5.68, 0.08, 1.38, ACCENT_GOLD)
    add_textbox(
        slide,
        7.00,
        5.78,
        5.55,
        0.34,
        [
            {
                "runs": [
                    {"text": "金茂大厦 56 层云端观景自助  ", "size": 14.5, "color": ATOMY_DARK, "bold": True},
                    {"text": "300元 / 位", "size": 13.5, "color": ACCENT_GOLD, "bold": True},
                ]
            }
        ],
    )
    add_textbox(
        slide,
        7.00,
        6.14,
        5.55,
        0.84,
        [
            {
                "text": "得益于公司的旅游补贴，我们特意带父母和妻儿登上了金茂大厦 56 层，体验了 300元/位 的高空观景自助餐。在壮丽的云端中庭与城市美景陪伴下，全家人品尝丰盛中西料理，父母连连夸赞这次体验既体面又难忘。",
                "size": 11.8,
                "color": TEXT_BODY,
                "line": 1.38,
            }
        ],
    )


# ===========================================================================
# 第 4 页：上海站（二）豫园与城隍庙
# ===========================================================================
def build_slide_4_shanghai_culture(prs, blank_layout, slot_photos, specs_by_id):
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_WARM)
    add_header(
        slide,
        "03 / 上海站 · 江南古韵",
        "游豫园亭台楼阁，品老城隍庙人间烟火",
        "在摩登都市中心寻访明代古典园林，带父母重温记忆中的老上海风情",
        "04 / 06",
    )

    # 左侧：上下或左右双图展示（豫园 + 城隍庙）
    insert_framed_photo(
        slide,
        0.73,
        1.98,
        4.10,
        4.55,
        slot_photos["yuyuan"],
        specs_by_id["yuyuan"],
        caption="▲ 豫园古典园林 · 一步一景的江南雅致",
    )

    insert_framed_photo(
        slide,
        5.05,
        1.98,
        4.10,
        4.55,
        slot_photos["chenghuang"],
        specs_by_id["chenghuang"],
        caption="▲ 老城隍庙街区 · 飞檐翘角与地道风味",
    )

    # 右侧：深度游览心得面板
    add_rect(slide, 9.40, 1.90, 3.28, 5.12, CARD_WHITE, BORDER_LINE, 1.0, rounded=True)
    add_rect(slide, 9.40, 1.90, 3.28, 0.10, ATOMY_BLUE)

    sections = [
        (
            "曲径通幽 · 豫园赏景",
            "走进拥有四百余年历史的豫园，假山叠石、池塘锦鲤与亭台楼阁相映成趣。父母最喜欢这样古色古香的江南园林，慢步在九曲桥畔，一边欣赏古建雕刻，一边给孩子讲述历史故事。",
        ),
        (
            "市井温情 · 城隍庙寻味",
            "毗邻豫园的老城隍庙热闹非凡，飞檐斗拱间飘散着各色传统糕点与小笼包的香气。带着家人品尝地道上海特色小吃，感受浓浓的市井烟火气。",
        ),
        (
            "慢下来的陪伴时光",
            "平时总是步履匆匆，而在豫园与城隍庙的半日闲游里，看着父母舒展的眉头、妻子温柔的照顾和孩子好奇的眼神，深感『陪伴』才是给家人最好的礼物。",
        ),
    ]

    for idx, (s_title, s_text) in enumerate(sections):
        sy = 2.18 + idx * 1.56
        add_textbox(
            slide,
            9.62,
            sy,
            2.84,
            0.34,
            [{"text": f"■  {s_title}", "size": 13.5, "color": ATOMY_BLUE if idx != 1 else ACCENT_GOLD, "bold": True}],
        )
        add_textbox(
            slide,
            9.62,
            sy + 0.34,
            2.84,
            1.12,
            [{"text": s_text, "size": 11.5, "color": TEXT_BODY, "line": 1.38}],
        )


# ===========================================================================
# 第 5 页：杭州站 —— 西湖、雷峰塔与海鲜大餐
# ===========================================================================
def build_slide_5_hangzhou(prs, blank_layout, slot_photos, specs_by_id):
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_MAIN)
    add_header(
        slide,
        "04 / 杭州站 · 山水与盛宴",
        "赏西湖烟雨登雷峰塔，享龙虾面包蟹海鲜大餐",
        "从『淡妆浓抹总相宜』的西子湖畔，到全家人大快朵颐的海鲜盛宴",
        "05 / 06",
    )

    # 左栏：西湖与雷峰塔
    insert_framed_photo(
        slide,
        0.73,
        1.98,
        5.75,
        3.15,
        slot_photos["westlake"],
        specs_by_id["westlake"],
        caption="▲ 杭州西湖与雷峰塔 · 登塔远眺一湖碧水与满城翠色",
    )

    add_rect(slide, 0.65, 5.68, 5.91, 1.38, CARD_WHITE, BORDER_LINE, 1.0, rounded=True)
    add_rect(slide, 0.65, 5.68, 0.08, 1.38, ATOMY_BLUE)
    add_textbox(
        slide,
        0.88,
        5.78,
        5.55,
        0.34,
        [{"text": "西湖漫步 & 雷峰塔登高望远", "size": 14.5, "color": ATOMY_DARK, "bold": True}],
    )
    add_textbox(
        slide,
        0.88,
        6.14,
        5.55,
        0.84,
        [
            {
                "text": "来到人间天堂杭州，带着父母、老婆和孩子漫步西湖岸边，杨柳依依、碧波荡漾。随后全家登上雷峰塔，俯瞰西湖全景与湖心亭台，父母圆了多年想看西湖与雷峰塔的心愿，孩子也实地见到了课本与传说中的名胜。",
                "size": 11.8,
                "color": TEXT_BODY,
                "line": 1.38,
            }
        ],
    )

    # 右栏：龙虾、面包蟹等海鲜大餐
    insert_framed_photo(
        slide,
        6.85,
        1.98,
        5.75,
        3.15,
        slot_photos["seafood"],
        specs_by_id["seafood"],
        caption="▲ 丰盛海鲜大餐（大龙虾、面包蟹等） · 全家人围坐共享美味",
    )

    add_rect(slide, 6.77, 5.68, 5.91, 1.38, CARD_WHITE, BORDER_LINE, 1.0, rounded=True)
    add_rect(slide, 6.77, 5.68, 0.08, 1.38, ACCENT_GOLD)
    add_textbox(
        slide,
        7.00,
        5.78,
        5.55,
        0.34,
        [
            {
                "runs": [
                    {"text": "犒劳全家的海鲜大餐  ", "size": 14.5, "color": ATOMY_DARK, "bold": True},
                    {"text": "龙虾 · 面包蟹 · 鲜活海味", "size": 13, "color": ACCENT_GOLD, "bold": True},
                ]
            }
        ],
    )
    add_textbox(
        slide,
        7.00,
        6.14,
        5.55,
        0.84,
        [
            {
                "text": "旅途中最幸福的时刻莫过于全家人围坐一桌共享美食！我们特意安排了丰盛的海鲜大餐，鲜嫩饱满的大龙虾、蟹膏满满的面包蟹摆满桌面，老人和孩子吃得津津有味，欢声笑语里满是幸福与满足。",
                "size": 11.8,
                "color": TEXT_BODY,
                "line": 1.38,
            }
        ],
    )


# ===========================================================================
# 第 6 页：旅游心得与感恩寄语
# ===========================================================================
def build_slide_6_reflection(prs, blank_layout, slot_photos, specs_by_id):
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_WARM)
    add_header(
        slide,
        "05 / 心得与感恩",
        "五年成长心怀感恩，化作动力再谱新篇",
        "一次有温度的员工福利之旅 —— 连接起个人奋斗、家庭幸福与企业关怀",
        "06 / 06",
    )

    # 左侧：三大核心感悟卡片
    reflections = [
        (
            "01",
            "一份福利，让家人共享成长的荣耀",
            "在艾多美工作的 5 年，是我个人能力不断历练的 5 年。而公司每满 3、5、7、10 年由公司承担费用的员工旅游福利，让我能把这份职场收获直接转化为父母、妻子和孩子的幸福体验。看到父母在金茂 56 层和西湖边的笑容，我为身在艾多美而倍感自豪。",
        ),
        (
            "02",
            "一次同行，拉近三代人的心与爱",
            "从上海外滩、南京路、豫园、城隍庙，到杭州西湖、雷峰塔，再到共享龙虾与面包蟹海鲜大餐，几天的朝夕相伴弥补了平日忙碌的亏欠。家人的理解与支持，始终是我全身心投入工作的最坚强后盾。",
        ),
        (
            "03",
            "一份期许，奔赴下一个 7 年与 10 年",
            "旅行虽已圆满收官，但感恩与干劲常留心间。未来我将带着旅途中积蓄的温暖与能量，在工作岗位上更加脚踏实地、全力以赴，与艾多美并肩同行，期待满 7 年、满 10 年时再带家人开启新的旅程！",
        ),
    ]

    for i, (num, r_title, r_body) in enumerate(reflections):
        ry = 1.95 + i * 1.65
        add_rect(slide, 0.65, ry, 7.20, 1.48, CARD_WHITE, BORDER_LINE, 1.0, rounded=True)
        # 数字徽章
        add_rect(slide, 0.88, ry + 0.22, 0.56, 0.56, ATOMY_BLUE if i != 1 else ACCENT_GOLD, rounded=True)
        add_textbox(
            slide,
            0.88,
            ry + 0.22,
            0.56,
            0.56,
            [{"text": num, "size": 16, "color": CARD_WHITE, "bold": True, "align": "center"}],
            valign="middle",
        )
        add_textbox(
            slide,
            1.62,
            ry + 0.16,
            6.05,
            0.34,
            [{"text": r_title, "size": 15, "color": ATOMY_DARK, "bold": True}],
        )
        add_textbox(
            slide,
            1.62,
            ry + 0.52,
            6.05,
            0.86,
            [{"text": r_body, "size": 11.5, "color": TEXT_BODY, "line": 1.36}],
        )

    # 右侧：温馨留念照 + 感恩寄语金句框
    insert_framed_photo(
        slide,
        8.15,
        1.98,
        4.45,
        3.25,
        slot_photos["closing"],
        specs_by_id["closing"],
        caption="▲ 旅途剪影 · 感恩公司赋予的美好家庭时光",
    )

    add_rect(slide, 8.07, 5.78, 4.61, 1.12, ATOMY_DARK, None, rounded=True)
    add_textbox(
        slide,
        8.28,
        5.88,
        4.20,
        0.92,
        [
            {"text": "“感恩艾多美，感恩家人的陪伴。", "size": 13.5, "color": CARD_WHITE, "bold": True, "align": "center"},
            {
                "text": "五年是里程碑，更是新起点，我们下一个五年再见！”",
                "size": 12.5,
                "color": "BAE6FD",
                "bold": True,
                "align": "center",
                "before": 8,
            },
        ],
        valign="middle",
    )


def build_presentation(photo_dir, output_path):
    photo_files, resolved_dir = discover_photos(photo_dir)
    if photo_files:
        print(f"已从『{resolved_dir}』加载 {len(photo_files)} 张旅游照片：")
        for pf in photo_files:
            print(f"  - {pf.name}")
    else:
        print(
            f"未在本地路径检测到照片文件（查找路径：{photo_dir}），\n"
            "已自动使用高清主题版式图生成 PPT；在本机运行脚本即可一键嵌入真实现场照片。"
        )

    slot_photos = match_photos_to_slots(photo_files)
    specs_by_id = {spec["id"]: spec for spec in SLOT_SPECS}

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    blank_layout = prs.slide_layouts[6]

    build_slide_1_cover(prs, blank_layout, slot_photos, specs_by_id)
    build_slide_2_welfare(prs, blank_layout, slot_photos, specs_by_id)
    build_slide_3_shanghai_modern(prs, blank_layout, slot_photos, specs_by_id)
    build_slide_4_shanghai_culture(prs, blank_layout, slot_photos, specs_by_id)
    build_slide_5_hangzhou(prs, blank_layout, slot_photos, specs_by_id)
    build_slide_6_reflection(prs, blank_layout, slot_photos, specs_by_id)

    prs.save(output_path)
    print(f"\n已成功生成 PPT：{output_path}（共 {len(prs.slides)} 页）")


def parse_args():
    parser = argparse.ArgumentParser(description="生成艾多美5周年员工旅游心得 PPT（自动配图）")
    parser.add_argument(
        "-i",
        "--input",
        default="/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片",
        help="照片文件夹路径（默认：/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片）",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=DEFAULT_OUTPUT,
        help=f"输出 PPTX 文件路径（默认：{DEFAULT_OUTPUT}）",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    build_presentation(args.input, args.output)
