#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
艾多美 5 周年员工旅游心得 PPT 生成器
====================================

根据艾多美满 3/5/7/10 年员工旅游福利（公司全额承担费用），
生成「满 5 周年 · 携父母、妻子、孩子共赴上海 & 杭州」的 6 页旅游心得 PPT。

自动读取指定照片文件夹中的照片（支持按文件名关键词智能匹配或按顺序自动分配，
并自动居中裁剪填充至照片框、修正手机 EXIF 方向、支持 JPG/PNG/HEIC/WebP）；
若在云端或未找到照片目录，将自动生成精美的主题插画占位图，并在页面底部标注替换提示。

用法：
  # 默认直接读取 /Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片
  python3 make_trip_ppt.py

  # 或手动指定照片目录与输出路径
  python3 make_trip_ppt.py -p "/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片" -o "艾多美5周年员工旅游心得_沪杭家庭之旅.pptx"
"""

import argparse
import io
import math
import os
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except ImportError:
    pass

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt


DEFAULT_PHOTO_DIR = "/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片"
DEFAULT_OUTPUT = "艾多美5周年员工旅游心得_沪杭家庭之旅.pptx"

# 字体与配色（艾多美蓝 + 暖金家居温情色系）
SANS = "PingFang SC"
SANS_FALLBACK = "Microsoft YaHei"

ATOMY_BLUE = "0093D0"       # 艾多美经典蓝
ATOMY_NAVY = "0B253A"       # 深海蓝（封面/尾页底色）
ATOMY_DARK_CARD = "13334C"  # 深蓝卡片
ATOMY_LIGHT = "EBF6FC"      # 浅蓝柔和底色
WARM_BG = "F8F9FA"          # 页面浅暖灰底色
CARD_WHITE = "FFFFFF"       # 纯白卡片
GOLD = "D9822B"             # 暖金色（5周年纪念点缀）
GOLD_LIGHT = "FDF5EB"       # 浅暖金底色
INK = "1A242D"              # 主标题深墨色
BODY = "3A4652"             # 正文深灰
MUTED = "7A8794"            # 辅助说明灰
BORDER = "E1E7EC"           # 细边框色

W = 13.333333  # 16:9 宽（英寸）
H = 7.5        # 16:9 高（英寸）
ML = 0.68      # 页面左边距


# ---------------------------------------------------------------------------
# 10 个照片槽位定义（用于智能匹配文件名或生成对应主题的精美插画卡片）
# ---------------------------------------------------------------------------
PHOTO_SLOTS = [
    {
        "slot_id": "sh_bund",
        "title": "上海外滩 · 万国建筑与江风",
        "subtitle": "携家人漫步黄浦江畔 · 留念东方明珠",
        "keywords": ["外滩", "东方明珠", "浦东", "黄浦江", "bund", "陆家嘴"],
        "theme": "bund",
        "priority": 1,
    },
    {
        "slot_id": "sh_jinmao",
        "title": "金茂大厦 56 层 · 云端观景自助",
        "subtitle": "300元/位 · 俯瞰申城天际线共享家庭盛宴",
        "keywords": ["金茂", "56", "自助", "观景", "云端", "jinmao", "大厦", "高空"],
        "theme": "jinmao",
        "priority": 1,
    },
    {
        "slot_id": "sh_yuyuan",
        "title": "豫园 & 城隍庙 · 江南古韵",
        "subtitle": "九曲桥畔 · 陪父母品味老上海慢时光",
        "keywords": ["豫园", "城隍庙", "九曲桥", "古镇", "园林", "小笼包", "yuyuan"],
        "theme": "yuyuan",
        "priority": 1,
    },
    {
        "slot_id": "sh_nanjing",
        "title": "南京路步行街 · 十里洋场繁华",
        "subtitle": "铛铛车与霓虹街景 · 孩子的欢乐天地",
        "keywords": ["南京路", "步行街", "南京东路", "当当车", "铛铛车", "nanjing"],
        "theme": "nanjing",
        "priority": 1,
    },
    {
        "slot_id": "hz_westlake",
        "title": "杭州西湖 · 湖光山色入画来",
        "subtitle": "淡妆浓抹总相宜 · 一家人泛舟漫步苏堤",
        "keywords": ["西湖", "断桥", "苏堤", "游船", "三潭印月", "westlake", "xihu"],
        "theme": "westlake",
        "priority": 1,
    },
    {
        "slot_id": "hz_leifeng",
        "title": "雷峰塔 · 登高远眺西湖全景",
        "subtitle": "凭栏揽胜 · 与父母孩子共赏千年古塔",
        "keywords": ["雷峰塔", "雷峰", "登塔", "夕照", "leifeng", "俯瞰西湖"],
        "theme": "leifeng",
        "priority": 1,
    },
    {
        "slot_id": "seafood_lobster",
        "title": "海鲜大餐 · 鲜美大龙虾",
        "subtitle": "虾肉紧实弹牙 · 犒劳全家的舌尖盛宴",
        "keywords": ["龙虾", "大虾", "虾", "lobster"],
        "theme": "lobster",
        "priority": 1,
    },
    {
        "slot_id": "seafood_crab",
        "title": "海鲜大餐 · 膏满黄肥面包蟹",
        "subtitle": "面包蟹与丰富海味 · 父母孩子赞不绝口",
        "keywords": ["面包蟹", "螃蟹", "蟹", "海鲜", "大餐", "美食", "crab", "seafood"],
        "theme": "crab",
        "priority": 2,
    },
    {
        "slot_id": "cover_family",
        "title": "全家福 · 沪杭同行",
        "subtitle": "父母 · 妻子 · 孩子 · 满5周年感恩之旅",
        "keywords": ["全家", "家庭", "合影", "家人", "出发", "父母", "老婆", "孩子", "封面", "family", "cover"],
        "theme": "family",
        "priority": 3,
    },
    {
        "slot_id": "welfare_happy",
        "title": "三代同堂 · 旅途温馨瞬间",
        "subtitle": "公司承担费用 · 安心带全家看世界",
        "keywords": ["父母", "爸妈", "老婆", "孩子", "高铁", "酒店", "开心", "笑容", "家庭", "合影", "上海", "杭州"],
        "theme": "warm",
        "priority": 4,
    },
]


# ---------------------------------------------------------------------------
# 照片扫描、智能匹配与居中裁剪处理
# ---------------------------------------------------------------------------
def collect_photos(dirs):
    """扫描候选目录，返回按文件名排序的图片路径列表。"""
    exts = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".heic", ".heif", ".tif", ".tiff"}
    found = []
    for d in dirs:
        if not d:
            continue
        p = Path(d).expanduser()
        if p.is_file() and p.suffix.lower() in exts:
            found.append(p)
        elif p.is_dir():
            for child in sorted(p.rglob("*")):
                if child.is_file() and child.suffix.lower() in exts and not child.name.startswith("."):
                    found.append(child)
            if found:
                break
    return found


def assign_photos_to_slots(photos):
    """
    将用户目录下的照片智能分配给 10 个展示位：
    1. 优先根据文件名中包含的景点/美食关键词匹配（如“外滩”、“金茂”、“西湖”、“龙虾”等）；
    2. 剩余未指定关键词的槽位，按照片顺序依次填入（若照片少于 10 张则循环复用）。
    """
    if not photos:
        return {slot["slot_id"]: None for slot in PHOTO_SLOTS}

    assigned = {}
    used_indices = set()

    # 第一轮：按文件名关键词匹配
    for slot in PHOTO_SLOTS:
        sid = slot["slot_id"]
        best_idx = None
        for idx, path in enumerate(photos):
            if idx in used_indices:
                continue
            stem = path.stem.lower()
            if any(kw.lower() in stem for kw in slot["keywords"]):
                best_idx = idx
                break
        if best_idx is not None:
            assigned[sid] = photos[best_idx]
            used_indices.add(best_idx)

    # 第二轮：剩余照片按顺序填入未分配的槽位
    remaining = [p for i, p in enumerate(photos) if i not in used_indices]
    rem_cursor = 0
    all_cursor = 0

    for slot in PHOTO_SLOTS:
        sid = slot["slot_id"]
        if sid in assigned:
            continue
        if rem_cursor < len(remaining):
            assigned[sid] = remaining[rem_cursor]
            rem_cursor += 1
        else:
            assigned[sid] = photos[all_cursor % len(photos)]
            all_cursor += 1

    return assigned


def _load_pil_font(size_px):
    """加载用于生成占位插画卡片的中文字体。"""
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
                continue
    return ImageFont.load_default()


def render_scenic_placeholder(slot_info, width_in, height_in, dpi=180):
    """
    当运行环境尚未挂载本地照片目录时，生成带有对应景点/美食矢量剪影风格的视觉图，
    确保生成的 PPT 美观完整，且用户在有照片的电脑上重跑脚本即可一键替换为真实照片。
    """
    w_px = max(320, round(width_in * dpi))
    h_px = max(220, round(height_in * dpi))
    theme = slot_info.get("theme", "family")

    palettes = {
        "family":   ((14, 52, 82),   (28, 108, 158), (242, 182, 98)),
        "warm":     ((26, 72, 105),  (58, 134, 178), (245, 200, 125)),
        "bund":     ((18, 38, 70),   (45, 85, 138),  (242, 190, 102)),
        "jinmao":   ((32, 48, 78),   (78, 106, 152), (246, 205, 120)),
        "yuyuan":   ((78, 36, 38),   (138, 66, 62),  (242, 202, 135)),
        "nanjing":  ((42, 34, 72),   (96, 72, 138),  (248, 198, 112)),
        "westlake": ((22, 68, 74),   (54, 128, 126), (218, 238, 226)),
        "leifeng":  ((58, 52, 42),   (126, 102, 72), (246, 206, 128)),
        "lobster":  ((92, 32, 28),   (168, 64, 52),  (250, 208, 132)),
        "crab":     ((88, 42, 24),   (165, 88, 44),  (252, 214, 142)),
    }
    c_top, c_bot, c_accent = palettes.get(theme, palettes["family"])

    img = Image.new("RGB", (w_px, h_px), c_top)
    draw = ImageDraw.Draw(img)

    # 绘制柔和渐变背景
    for y in range(h_px):
        t = y / max(1, h_px - 1)
        r = int(c_top[0] * (1 - t) + c_bot[0] * t)
        g = int(c_top[1] * (1 - t) + c_bot[1] * t)
        b = int(c_top[2] * (1 - t) + c_bot[2] * t)
        draw.line([(0, y), (w_px, y)], fill=(r, g, b))

    # 绘制主题意境几何装饰
    cx, cy = w_px // 2, int(h_px * 0.42)
    r_sun = int(min(w_px, h_px) * 0.16)
    draw.ellipse(
        [cx - r_sun, cy - r_sun, cx + r_sun, cy + r_sun],
        fill=c_accent,
    )

    if theme in ("bund", "jinmao", "nanjing"):
        # 城市天际线剪影（东方明珠 / 金茂大厦风格）
        base_y = int(h_px * 0.72)
        buildings = [
            (0.12, 0.07, 0.22), (0.21, 0.06, 0.30), (0.30, 0.08, 0.18),
            (0.42, 0.04, 0.44), (0.50, 0.07, 0.38), (0.60, 0.08, 0.42),
            (0.71, 0.06, 0.26), (0.80, 0.08, 0.20),
        ]
        for bx_r, bw_r, bh_r in buildings:
            x0 = int(w_px * bx_r)
            x1 = int(w_px * (bx_r + bw_r))
            y0 = base_y - int(h_px * bh_r)
            draw.rectangle([x0, y0, x1, base_y], fill=c_top)
        # 东方明珠球体点缀
        tower_x = int(w_px * 0.44)
        draw.line([(tower_x, base_y - int(h_px * 0.50)), (tower_x, base_y)], fill=c_accent, width=3)
        draw.ellipse([tower_x - 10, base_y - int(h_px * 0.36) - 10, tower_x + 10, base_y - int(h_px * 0.36) + 10], fill=c_accent)
        draw.rectangle([0, base_y, w_px, h_px], fill=(c_top[0] // 2, c_top[1] // 2, c_top[2] // 2))
    elif theme in ("westlake", "leifeng", "yuyuan"):
        # 江南远山与湖水波纹 / 古塔屋檐
        base_y = int(h_px * 0.72)
        pts_mtn1 = [(0, base_y), (int(w_px * 0.28), int(h_px * 0.46)), (int(w_px * 0.62), base_y)]
        pts_mtn2 = [(int(w_px * 0.35), base_y), (int(w_px * 0.72), int(h_px * 0.42)), (w_px, base_y)]
        draw.polygon(pts_mtn1, fill=(c_top[0] + 10, c_top[1] + 18, c_top[2] + 18))
        draw.polygon(pts_mtn2, fill=c_top)
        draw.rectangle([0, base_y, w_px, h_px], fill=(c_top[0] // 2, c_top[1] // 2, c_top[2] // 2))
    else:
        # 温馨光晕与底部深色衬底
        base_y = int(h_px * 0.70)
        draw.rectangle([0, base_y, w_px, h_px], fill=(c_top[0] // 2, c_top[1] // 2, c_top[2] // 2))

    # 边框装饰线
    pad = 12
    draw.rectangle([pad, pad, w_px - pad, h_px - pad], outline=c_accent, width=2)

    # 绘制文字标题与副标题
    f_title = _load_pil_font(max(18, int(min(w_px, h_px) * 0.072)))
    f_sub = _load_pil_font(max(13, int(min(w_px, h_px) * 0.045)))

    title_text = slot_info["title"]
    sub_text = slot_info["subtitle"]

    tb_box = draw.textbbox((0, 0), title_text, font=f_title)
    tw = tb_box[2] - tb_box[0]
    draw.text(((w_px - tw) / 2, int(h_px * 0.74)), title_text, font=f_title, fill=(255, 255, 255))

    sb_box = draw.textbbox((0, 0), sub_text, font=f_sub)
    sw = sb_box[2] - sb_box[0]
    draw.text(((w_px - sw) / 2, int(h_px * 0.85)), sub_text, font=f_sub, fill=c_accent)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def crop_photo_to_box(photo_path, width_in, height_in, dpi=220):
    """读取用户真实照片，修正 EXIF 旋转，按目标宽高比居中裁剪并返回 PNG 字节流。"""
    img = Image.open(photo_path)
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGB")

    target_w = max(200, round(width_in * dpi))
    target_h = max(200, round(height_in * dpi))
    fitted = ImageOps.fit(img, (target_w, target_h), method=Image.Resampling.LANCZOS, centering=(0.5, 0.5))

    buf = io.BytesIO()
    fitted.save(buf, format="JPEG", quality=92)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# PPTX 基础排版绘制组件
# ---------------------------------------------------------------------------
def rgb(hex_str):
    return RGBColor.from_string(hex_str)


def _anchor(tf, valign):
    body = tf._txBody.find(qn("a:bodyPr"))
    body.set("anchor", {"top": "t", "middle": "ctr", "bottom": "b"}[valign])
    for key in ("lIns", "rIns", "tIns", "bIns"):
        body.set(key, "0")


def _run(p, text, font, size, color, bold=False):
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = False
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


def tb(slide, x, y, w, h, paragraphs, valign="top"):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    _anchor(tf, valign)
    aligns = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    for i, info in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = aligns[info.get("align", "left")]
        if "line" in info:
            p.line_spacing = info["line"]
        if info.get("before"):
            p.space_before = Pt(info["before"])
        if info.get("after"):
            p.space_after = Pt(info["after"])
        runs = info.get("runs")
        if runs:
            for item in runs:
                _run(
                    p,
                    item["text"],
                    item.get("font", SANS),
                    item["size"],
                    item["color"],
                    item.get("bold", False),
                )
        else:
            _run(
                p,
                info["text"],
                info.get("font", SANS),
                info["size"],
                info["color"],
                info.get("bold", False),
            )
    return shape


def rect(slide, x, y, w, h, fill, line_color=None, line_pt=0.75):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    if line_color:
        shape.line.color.rgb = rgb(line_color)
        shape.line.width = Pt(line_pt)
    else:
        shape.line.fill.background()
    return shape


def bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color)


def new_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def add_header(slide, step_num, tag, title):
    """统一页面顶部栏：艾多美蓝标签 + 大标题 + 细分割线。"""
    rect(slide, 0, 0, W, 0.06, ATOMY_BLUE)
    tb(
        slide, ML, 0.32, 9.5, 0.28,
        [{"text": f"{step_num}  ·  {tag}", "size": 12.5, "color": ATOMY_BLUE, "bold": True}],
    )
    tb(
        slide, ML, 0.62, 11.9, 0.46,
        [{"text": title, "size": 25, "color": INK, "bold": True}],
    )
    rect(slide, ML, 1.18, 11.97, 0.015, BORDER)


def add_footer(slide, page_num, total_pages=6, dark=False):
    """统一页脚与页码。"""
    left_color = "7E97AD" if dark else MUTED
    tb(
        slide, ML, 7.06, 8.5, 0.26,
        [{"text": "艾多美（Atomy）员工 5 周年感恩旅游心得  ·  上海 & 杭州家庭之旅", "size": 10.5, "color": left_color}],
    )
    tb(
        slide, 11.2, 7.06, 1.45, 0.26,
        [{"text": f"{page_num:02d} / {total_pages:02d}", "size": 10.5, "color": left_color, "align": "right", "bold": True}],
    )


def add_photo_card(slide, x, y, w, h, slot_id, slot_map, caption=None, frame_bg=CARD_WHITE):
    """在指定区域插入带相框与底部图注的照片。"""
    slot_info = next(s for s in PHOTO_SLOTS if s["slot_id"] == slot_id)
    photo_path = slot_map.get(slot_id)

    # 白色相框底衬
    rect(slide, x, y, w, h, frame_bg, BORDER)
    pad = 0.08
    cap_h = 0.38 if caption else 0.0
    img_x = x + pad
    img_y = y + pad
    img_w = w - pad * 2
    img_h = h - pad * 2 - cap_h

    if photo_path and os.path.exists(photo_path):
        try:
            stream = crop_photo_to_box(photo_path, img_w, img_h)
        except Exception:
            stream = render_scenic_placeholder(slot_info, img_w, img_h)
    else:
        stream = render_scenic_placeholder(slot_info, img_w, img_h)

    slide.shapes.add_picture(stream, Inches(img_x), Inches(img_y), Inches(img_w), Inches(img_h))

    if caption:
        tb(
            slide, x + 0.14, y + h - cap_h - 0.02, w - 0.28, cap_h,
            [{"text": caption, "size": 11.5, "color": BODY, "align": "center", "bold": True}],
            valign="middle",
        )


# ---------------------------------------------------------------------------
# 构建 6 页 PPT 内容
# ---------------------------------------------------------------------------
def build_presentation(photo_dir, extra_dirs=None):
    candidate_dirs = [photo_dir]
    if extra_dirs:
        candidate_dirs.extend(extra_dirs)
    # 同时也支持当前目录下的 photos 文件夹
    candidate_dirs.append(str(Path(__file__).resolve().parent / "photos"))
    candidate_dirs.append(str(Path(__file__).resolve().parent / "5周年旅游照片"))

    photos = collect_photos(candidate_dirs)
    slot_map = assign_photos_to_slots(photos)

    prs = Presentation()
    prs.slide_width = Emu(12192000)   # 13.333 in
    prs.slide_height = Emu(6858000)   # 7.5 in
    prs.core_properties.title = "感恩五年相伴 · 携爱共赴沪杭——艾多美5周年员工旅游心得"
    prs.core_properties.subject = "艾多美员工满3/5/7/10年福利旅游心得分享"

    total_pages = 6

    # =========================================================================
    # 第 1 页：封面 —— 感恩五年相伴 · 携爱共赴沪杭
    # =========================================================================
    s1 = new_slide(prs)
    bg(s1, ATOMY_NAVY)
    rect(s1, 0, 0, 0.10, H, ATOMY_BLUE)
    rect(s1, 0.10, 0, 0.03, H, GOLD)

    # 左半区：标题与行程亮点卡片
    tb(
        s1, 0.78, 0.78, 7.2, 0.34,
        [{"text": "ATOMY 艾多美  ·  满 5 周年员工福利旅游心得分享", "size": 13.5, "color": "6EC3EB", "bold": True}],
    )
    tb(
        s1, 0.78, 1.32, 7.3, 1.85,
        [
            {"text": "感恩五年相伴", "size": 42, "color": CARD_WHITE, "bold": True, "after": 6},
            {"text": "携爱共赴沪杭之旅", "size": 40, "color": GOLD, "bold": True},
        ],
    )
    rect(s1, 0.78, 3.30, 1.35, 0.045, ATOMY_BLUE)

    tb(
        s1, 0.78, 3.55, 7.1, 1.15,
        [
            {
                "text": "员工每满 3、5、7、10 年专享员工旅游  ·  费用由公司全额承担",
                "size": 16,
                "color": CARD_WHITE,
                "bold": True,
                "after": 8,
            },
            {
                "text": "带着父母、老婆与孩子，一家人走过上海的繁华天际与杭州的湖山诗意，\n把五年奋斗的收获，化作三代同堂最温暖的笑容。",
                "size": 13.8,
                "color": "BDD4E6",
                "line": 1.38,
            },
        ],
    )

    badges = [
        ("同行家人", "父母 · 老婆 · 孩子"),
        ("第一站 · 上海", "外滩 · 金茂56层 · 豫园 · 南京路"),
        ("第二站 · 杭州", "西湖泛舟 · 雷峰塔登高"),
        ("舌尖盛宴", "金茂观景自助 · 龙虾面包蟹大餐"),
    ]
    bx = 0.78
    for b_label, b_val in badges:
        rect(s1, bx, 5.18, 1.72, 1.25, ATOMY_DARK_CARD, "234E70")
        rect(s1, bx, 5.18, 1.72, 0.04, GOLD)
        tb(s1, bx + 0.12, 5.34, 1.48, 0.28, [{"text": b_label, "size": 11.5, "color": "7CC7EB", "bold": True}])
        tb(s1, bx + 0.12, 5.66, 1.48, 0.65, [{"text": b_val, "size": 11.5, "color": CARD_WHITE, "line": 1.28}])
        bx += 1.84

    # 右半区：封面主照片（全家福）
    add_photo_card(
        s1, 8.28, 0.78, 4.38, 5.65,
        "cover_family", slot_map,
        caption="▲ 满 5 周年员工旅游 · 全家沪杭同行合影",
        frame_bg=CARD_WHITE,
    )
    add_footer(s1, 1, total_pages, dark=True)

    # =========================================================================
    # 第 2 页：缘起与感恩 —— 3/5/7/10 年福利制度与三代同堂出发
    # =========================================================================
    s2 = new_slide(prs)
    bg(s2, WARM_BG)
    add_header(s2, "01", "福利与初心", "公司全额承担的 5 周年礼物：带着父母、老婆、孩子一起看世界")

    # 顶部里程碑时间轴：3年、5年（高亮）、7年、10年
    milestones = [
        ("满 3 周年", "初绽芳华 · 感恩启程", False),
        ("满 5 周年（本次旅程）", "五载同心 · 携全家共游沪杭", True),
        ("满 7 周年", "笃行致远 · 再赴山海", False),
        ("满 10 周年", "十年树木 · 荣耀同行", False),
    ]
    mx = ML
    for m_year, m_desc, is_current in milestones:
        fill_c = ATOMY_BLUE if is_current else CARD_WHITE
        border_c = GOLD if is_current else BORDER
        title_c = CARD_WHITE if is_current else ATOMY_BLUE
        desc_c = "FFF3E0" if is_current else BODY
        rect(s2, mx, 1.38, 2.84, 1.02, fill_c, border_c, line_pt=1.5 if is_current else 0.75)
        if is_current:
            rect(s2, mx, 1.38, 2.84, 0.06, GOLD)
        tb(s2, mx + 0.18, 1.50, 2.48, 0.34, [{"text": m_year, "size": 15, "color": title_c, "bold": True}])
        tb(s2, mx + 0.18, 1.88, 2.48, 0.36, [{"text": m_desc, "size": 12, "color": desc_c}])
        mx += 3.04

    # 左侧：感悟与行程总览卡片
    rect(s2, ML, 2.60, 6.85, 4.18, CARD_WHITE, BORDER)
    rect(s2, ML, 2.60, 0.07, 4.18, ATOMY_BLUE)
    tb(
        s2, ML + 0.30, 2.80, 6.3, 0.36,
        [{"text": "为什么这次 5 周年员工旅游格外有意义？", "size": 17.5, "color": INK, "bold": True}],
    )

    points_s2 = [
        (
            "公司承担费用，让“陪伴家人”毫无后顾之忧",
            "在艾多美，员工每满 3、5、7、10 年即可享受一次公司承担费用的员工旅游。这份沉甸甸的福利不仅是物质奖励，更是公司对员工背后整个家庭的关怀。",
        ),
        (
            "敬父母、爱妻儿：三代同堂的圆梦之旅",
            "平时忙于工作，很少有机会带全家人一起出远门。借着满 5 周年的契机，特意带上爸妈、老婆和孩子——父母看风景、孩子长见识、爱人享轻松。",
        ),
        (
            "双城经典路线：上海摩登繁华 × 杭州诗画山水",
            "精心规划“上海 + 杭州”双城路线：既在魔都体验外滩、金茂大厦 56 层自助、豫园城隍庙与南京路，又赴杭州饱览西湖与雷峰塔，并共享丰盛的海鲜大餐。",
        ),
    ]
    py = 3.28
    for p_title, p_desc in points_s2:
        tb(s2, ML + 0.30, py, 6.3, 0.30, [{"text": f"●  {p_title}", "size": 14, "color": ATOMY_BLUE, "bold": True}])
        tb(s2, ML + 0.52, py + 0.30, 6.05, 0.72, [{"text": p_desc, "size": 12.5, "color": BODY, "line": 1.32}])
        py += 1.12

    # 右侧：家庭温馨合影
    add_photo_card(
        s2, 7.78, 2.60, 4.87, 4.18,
        "welfare_happy", slot_map,
        caption="▲ 携父母、妻子与孩子踏上 5 周年感恩之旅",
    )
    add_footer(s2, 2, total_pages)

    # =========================================================================
    # 第 3 页：上海站（上）—— 外滩风云与金茂大厦 56 层云端观景自助
    # =========================================================================
    s3 = new_slide(prs)
    bg(s3, WARM_BG)
    add_header(s3, "02", "第一站 · 魔都上海（上）", "漫步外滩万国建筑群，登金茂大厦 56 层享云端观景自助")

    # 左栏：上海外滩
    rect(s3, ML, 1.38, 5.82, 5.42, CARD_WHITE, BORDER)
    rect(s3, ML, 1.38, 5.82, 0.06, ATOMY_BLUE)
    add_photo_card(
        s3, ML + 0.20, 1.60, 5.42, 3.42,
        "sh_bund", slot_map,
        caption="▲ 上海外滩 · 与父母妻儿吹着江风看陆家嘴天际线",
    )
    tb(
        s3, ML + 0.24, 5.14, 5.34, 0.34,
        [{"text": "外滩漫步  ·  百年万国建筑与黄浦江盛景", "size": 16.5, "color": INK, "bold": True}],
    )
    tb(
        s3, ML + 0.24, 5.52, 5.34, 1.12,
        [{
            "text": "陪着父母与老婆孩子漫步在外滩江畔，一边是厚重的百年万国建筑博览群，对岸是高耸入云的东方明珠与陆家嘴摩天大楼。江风拂面，父母感叹时代发展的日新月异，孩子兴奋地指着江上的游轮，全家人定格下难忘的合影。",
            "size": 12.8, "color": BODY, "line": 1.36,
        }],
    )

    # 右栏：金茂大厦 56 层 300元/位 观景自助
    rect(s3, 6.83, 1.38, 5.82, 5.42, CARD_WHITE, BORDER)
    rect(s3, 6.83, 1.38, 5.82, 0.06, GOLD)
    add_photo_card(
        s3, 7.03, 1.60, 5.42, 3.42,
        "sh_jinmao", slot_map,
        caption="▲ 金茂大厦 56 层天庭观景自助（300 元/位）",
    )
    tb(
        s3, 7.07, 5.14, 5.34, 0.34,
        [{"text": "金茂大厦 56 层  ·  300 元/位云端观景自助餐", "size": 16.5, "color": GOLD, "bold": True}],
    )
    tb(
        s3, 7.07, 5.52, 5.34, 1.12,
        [{
            "text": "特意带全家人登上陆家嘴地标——金茂大厦 56 层，体验 300 元/位的高空观景自助餐。置身震撼的巨型环廊中庭，脚下是繁华的申城风光，眼前是琳琅满目的中西美食与甜品。看着爸妈和老婆孩子边赏云端美景边开心品尝美食，幸福感油然而生。",
            "size": 12.8, "color": BODY, "line": 1.36,
        }],
    )

    add_footer(s3, 3, total_pages)

    # =========================================================================
    # 第 4 页：上海站（下）—— 豫园、城隍庙古韵与南京路步行街繁华
    # =========================================================================
    s4 = new_slide(prs)
    bg(s4, WARM_BG)
    add_header(s4, "03", "第一站 · 魔都上海（下）", "寻味豫园与城隍庙老城厢烟火，感受南京路“十里洋场”繁华")

    # 左栏：豫园 & 城隍庙
    rect(s4, ML, 1.38, 5.82, 5.42, CARD_WHITE, BORDER)
    rect(s4, ML, 1.38, 5.82, 0.06, ATOMY_BLUE)
    add_photo_card(
        s4, ML + 0.20, 1.60, 5.42, 3.42,
        "sh_yuyuan", slot_map,
        caption="▲ 豫园与城隍庙 · 飞檐翘角与九曲桥畔的江南韵味",
    )
    tb(
        s4, ML + 0.24, 5.14, 5.34, 0.34,
        [{"text": "豫园 & 城隍庙  ·  江南古典园林与地道老上海风情", "size": 16.5, "color": INK, "bold": True}],
    )
    tb(
        s4, ML + 0.24, 5.52, 5.34, 1.12,
        [{
            "text": "走进豫园与城隍庙，飞檐翘角、亭台楼阁尽显明清江南园林的雅致。陪着父母走过九曲桥，寓意平平安安、顺遂吉祥；在热闹的城隍庙街区品尝地道上海特色点心，感受摩天都市里最具人情味的老城厢烟火气。",
            "size": 12.8, "color": BODY, "line": 1.36,
        }],
    )

    # 右栏：南京路步行街
    rect(s4, 6.83, 1.38, 5.82, 5.42, CARD_WHITE, BORDER)
    rect(s4, 6.83, 1.38, 5.82, 0.06, ATOMY_BLUE)
    add_photo_card(
        s4, 7.03, 1.60, 5.42, 3.42,
        "sh_nanjing", slot_map,
        caption="▲ 南京路步行街 · 百年商业街的璀璨活力",
    )
    tb(
        s4, 7.07, 5.14, 5.34, 0.34,
        [{"text": "南京路步行街  ·  “中华商业第一街”的摩登脉搏", "size": 16.5, "color": INK, "bold": True}],
    )
    tb(
        s4, 7.07, 5.52, 5.34, 1.12,
        [{
            "text": "从外滩一路逛到热闹非凡的南京路步行街，百年老字号与现代潮流名店交相辉映，复古铛铛车叮当作响穿行其间。老婆和孩子挑选着心仪的纪念品，全家人边走边逛，沉浸在十里洋场独有的繁华与欢乐之中。",
            "size": 12.8, "color": BODY, "line": 1.36,
        }],
    )

    add_footer(s4, 4, total_pages)

    # =========================================================================
    # 第 5 页：杭州站 —— 西湖烟雨诗画与雷峰塔登高揽胜
    # =========================================================================
    s5 = new_slide(prs)
    bg(s5, WARM_BG)
    add_header(s5, "04", "第二站 · 诗画杭州", "漫步西湖赏“淡妆浓抹总相宜”，登临雷峰塔俯瞰湖山全景")

    # 左栏：西湖
    rect(s5, ML, 1.38, 5.82, 5.42, CARD_WHITE, BORDER)
    rect(s5, ML, 1.38, 5.82, 0.06, ATOMY_BLUE)
    add_photo_card(
        s5, ML + 0.20, 1.60, 5.42, 3.42,
        "hz_westlake", slot_map,
        caption="▲ 杭州西湖 · 碧波荡漾与垂柳依依的江南画卷",
    )
    tb(
        s5, ML + 0.24, 5.14, 5.34, 0.34,
        [{"text": "杭州西湖  ·  水光潋滟晴方好，山色空蒙雨亦奇", "size": 16.5, "color": INK, "bold": True}],
    )
    tb(
        s5, ML + 0.24, 5.52, 5.34, 1.12,
        [{
            "text": "告别上海的都市节奏，一家人来到魂牵梦萦的杭州西湖。沿着湖畔漫步，看碧波荡漾、垂柳拂岸、游船点点。这正是父母一直心心念念的江南名胜，慢节奏的湖光山色让全家人身心彻底放松，孩子也在课本里的风景前流连忘返。",
            "size": 12.8, "color": BODY, "line": 1.36,
        }],
    )

    # 右栏：雷峰塔
    rect(s5, 6.83, 1.38, 5.82, 5.42, CARD_WHITE, BORDER)
    rect(s5, 6.83, 1.38, 5.82, 0.06, GOLD)
    add_photo_card(
        s5, 7.03, 1.60, 5.42, 3.42,
        "hz_leifeng", slot_map,
        caption="▲ 雷峰塔 · 登塔远眺西湖全景与千年人文古迹",
    )
    tb(
        s5, 7.07, 5.14, 5.34, 0.34,
        [{"text": "雷峰塔登高  ·  一眼望尽西湖山水与杭州城景", "size": 16.5, "color": GOLD, "bold": True}],
    )
    tb(
        s5, 7.07, 5.52, 5.34, 1.12,
        [{
            "text": "贴心的电梯设计让父母和孩子也能轻松登上雷峰塔顶。凭栏远眺，整个西湖的苏堤、三潭印月与远山近水尽收眼底；塔内珍贵的遗址与精美壁画更让孩子大开眼界。站在高处迎风眺望，一家人共享这份开阔与喜悦。",
            "size": 12.8, "color": BODY, "line": 1.36,
        }],
    )

    add_footer(s5, 5, total_pages)

    # =========================================================================
    # 第 6 页：舌尖海鲜盛宴与五年感恩结语 —— 龙虾、面包蟹与未来期许
    # =========================================================================
    s6 = new_slide(prs)
    bg(s6, WARM_BG)
    add_header(s6, "05", "舌尖盛宴 · 感恩结语", "龙虾与面包蟹海鲜大餐犒劳家人，感恩艾多美赋予的幸福力量")

    # 左侧双图：龙虾 + 面包蟹海鲜大餐
    add_photo_card(
        s6, ML, 1.38, 3.55, 2.62,
        "seafood_lobster", slot_map,
        caption="▲ 鲜活大龙虾 · 鲜甜弹牙",
    )
    add_photo_card(
        s6, 4.40, 1.38, 3.55, 2.62,
        "seafood_crab", slot_map,
        caption="▲ 膏满黄肥面包蟹 · 满桌海味",
    )

    # 左下：海鲜大餐文字卡片
    rect(s6, ML, 4.16, 7.27, 2.64, CARD_WHITE, BORDER)
    rect(s6, ML, 4.16, 0.07, 2.64, GOLD)
    tb(
        s6, ML + 0.28, 4.32, 6.8, 0.34,
        [{"text": "舌尖上的幸福欢聚  ·  大龙虾、面包蟹海鲜大餐", "size": 16.5, "color": GOLD, "bold": True}],
    )
    tb(
        s6, ML + 0.28, 4.72, 6.8, 1.90,
        [
            {
                "text": "除了金茂大厦 56 层 300 元/位的云端自助，旅途中最让全家人回味无穷的，还有丰盛的海鲜大餐！个头十足的大龙虾肉质紧实鲜甜，膏满黄肥的面包蟹香气四溢，再加上各式时令鲜活海味摆满一桌。",
                "size": 12.8, "color": BODY, "line": 1.36, "after": 8,
            },
            {
                "text": "看着父母笑得合不拢嘴、老婆和孩子吃得津津有味，举杯相庆的那一刻，真切感受到：努力工作最大的意义，就是让所爱之人拥有稳稳的幸福。",
                "size": 12.8, "color": INK, "bold": True, "line": 1.36,
            },
        ],
    )

    # 右侧深蓝卡片：5 周年心得感悟与未来寄语（7周年/10周年）
    rect(s6, 8.15, 1.38, 4.50, 5.42, ATOMY_NAVY)
    rect(s6, 8.15, 1.38, 4.50, 0.07, ATOMY_BLUE)
    tb(
        s6, 8.43, 1.66, 3.94, 0.32,
        [{"text": "FIVE-YEAR REFLECTION", "size": 11, "color": "6EC3EB", "bold": True}],
    )
    tb(
        s6, 8.43, 1.98, 3.94, 0.42,
        [{"text": "满 5 周年员工旅游心得感悟", "size": 20, "color": CARD_WHITE, "bold": True}],
    )
    rect(s6, 8.43, 2.50, 0.90, 0.035, GOLD)

    reflections = [
        ("01  感恩公司大爱福利", "满 3、5、7、10 年的免费员工旅游，真正体现了艾多美“珍视灵魂”与关爱员工家庭的企业温度。"),
        ("02  家人是最坚实的后盾", "五年的成长离不开父母、老婆和孩子的理解与支持，这场沪杭之旅是对家人最好的馈赠。"),
        ("03  满电归来，奔赴 7 年与 10 年", "带着旅途积蓄的温暖与能量回到工作岗位，以更饱满的热情全心投入，期待满 7 年、10 年再携家人共赴新山海！"),
    ]
    ry = 2.72
    for r_head, r_body in reflections:
        rect(s6, 8.43, ry, 3.94, 1.18, ATOMY_DARK_CARD, "234E70")
        tb(s6, 8.60, ry + 0.12, 3.60, 0.28, [{"text": r_head, "size": 13, "color": GOLD, "bold": True}])
        tb(s6, 8.60, ry + 0.42, 3.60, 0.68, [{"text": r_body, "size": 11.5, "color": "D6E6F2", "line": 1.28}])
        ry += 1.30

    add_footer(s6, 6, total_pages)

    return prs, len(photos), slot_map


def parse_args():
    p = argparse.ArgumentParser(description="生成艾多美 5 周年员工旅游心得 PPT")
    p.add_argument(
        "-p", "--photo-dir",
        default=DEFAULT_PHOTO_DIR,
        help=f"照片目录路径（默认：{DEFAULT_PHOTO_DIR}）",
    )
    p.add_argument(
        "-o", "--output",
        default=DEFAULT_OUTPUT,
        help=f"输出 PPTX 文件路径（默认：{DEFAULT_OUTPUT}）",
    )
    return p.parse_args()


def main():
    args = parse_args()
    prs, photo_count, slot_map = build_presentation(args.photo_dir)
    out_path = Path(args.output).expanduser().resolve()
    prs.save(str(out_path))

    print(f"✅ 已生成 PPT：{out_path}")
    if photo_count > 0:
        print(f"📸 已自动载入并排版 {photo_count} 张真实旅游照片（来自 {args.photo_dir}）：")
        for s in PHOTO_SLOTS:
            sid = s["slot_id"]
            p_file = slot_map.get(sid)
            print(f"   - [{s['title']}] <- {p_file.name if p_file else '默认插画'}")
    else:
        print(f"ℹ️ 当前运行环境未检测到本地路径：{args.photo_dir}")
        print("   已使用高颜值景点/美食主题卡片完成 6 页排版。")
        print("   👉 在您的 Mac 上只需直接运行：`python3 make_trip_ppt.py`，即可自动把该文件夹下的照片一键排入 PPT！")


if __name__ == "__main__":
    main()
