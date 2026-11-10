#!/usr/bin/env python3
"""
艾多美5周年员工旅游心得 PPT 生成脚本
用法：
  python3 generate_travel_ppt.py --photos "/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片"
"""

from __future__ import annotations

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

DEFAULT_PHOTO_DIR = "/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片"
DEFAULT_OUTPUT = "艾多美5周年员工旅游心得_上海杭州.pptx"

FONT_FAMILY = "PingFang SC"

# 配色定义（艾多美清新蓝 + 温暖金）
BG_LIGHT = "F6F8FB"
BG_DARK = "0F2C4C"
WHITE = "FFFFFF"
ATOMY_BLUE = "0078B5"
ATOMY_DARK = "0D3B66"
ATOMY_SOFT = "E6F3FA"
GOLD = "D98A2B"
GOLD_SOFT = "FDF3E6"
INK = "1D2733"
BODY = "425265"
MUTED = "7B8A9C"
BORDER = "DCE4EE"

SLIDE_W = 13.333333
SLIDE_H = 7.5

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}

# 10个照片槽位定义：(槽位key, 标题, 副标题, 匹配关键词列表, 占位图渐变主色)
PHOTO_SLOTS: List[Tuple[str, str, str, List[str], Tuple[int, int, int]]] = [
    (
        "cover_family",
        "全家福 · 5周年纪念",
        "携父母、爱人与孩子共赴沪杭之旅",
        ["全家", "合影", "家人", "父母", "老婆", "孩子", "封面", "family", "cover"],
        (13, 59, 102),
    ),
    (
        "bund",
        "上海外滩 · 万国建筑与江景",
        "漫步黄浦江畔，感受魔都璀璨风光",
        ["外滩", "东方明珠", "浦东", "黄浦江", "bund", "陆家嘴"],
        (0, 102, 160),
    ),
    (
        "jinmao",
        "金茂大厦56层 · 云端观景自助",
        "300元/位高空美食体验，俯瞰申城天际线",
        ["金茂", "56", "自助", "云端", "观景", "餐厅", "jinmao", "buffet"],
        (180, 110, 35),
    ),
    (
        "yuyuan",
        "上海豫园 · 江南古典园林",
        "亭台楼阁、小桥流水，父母流连忘返",
        ["豫园", "园林", "九曲桥", "yuyuan"],
        (42, 110, 95),
    ),
    (
        "chenghuangmiao",
        "老城隍庙 · 海派民俗与美食",
        "飞檐翘角间品尝地道小吃，其乐融融",
        ["城隍庙", "小吃", "南翔", "小笼", "庙"],
        (165, 68, 52),
    ),
    (
        "nanjinglu",
        "南京路步行街 · 十里洋场繁华",
        "老字号与现代潮流交汇，一家人欢声笑语",
        ["南京路", "步行街", "南京东路", "当当车", "上海街景", "上海"],
        (35, 85, 135),
    ),
    (
        "xihu",
        "杭州西湖 · 水光潋滟晴方好",
        "漫步苏堤湖畔，领略江南诗画山水",
        ["西湖", "断桥", "苏堤", "三潭印月", "游船", "xihu", "westlake"],
        (28, 115, 112),
    ),
    (
        "leifengta",
        "雷峰塔 · 登高凭栏揽西湖",
        "登塔远眺西湖全景，聆听千年人文故事",
        ["雷峰塔", "塔", "西湖全景", "杭州", "leifeng"],
        (140, 92, 42),
    ),
    (
        "lobster",
        "海鲜盛宴 · 鲜美大龙虾",
        "公司福利加持，与家人共享丰盛大龙虾",
        ["龙虾", "大龙虾", "海鲜", "大餐", "lobster", "seafood"],
        (185, 65, 45),
    ),
    (
        "crab",
        "膏满黄肥 · 面包蟹与海味",
        "面包蟹饱满鲜甜，老人孩子赞不绝口",
        ["面包蟹", "螃蟹", "蟹", "海味", "聚餐", "美食", "crab"],
        (195, 105, 38),
    ),
]


def rgb(hex_str: str) -> RGBColor:
    return RGBColor.from_string(hex_str)


def _load_cjk_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
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
                return ImageFont.truetype(path, size=size)
            except Exception:
                continue
    return ImageFont.load_default()


def discover_and_assign_photos(photo_dir: Path) -> Tuple[Dict[str, Optional[Path]], List[Path]]:
    """扫描照片目录，优先按文件名关键词匹配各景点槽位，剩余照片按顺序自动补齐。"""
    all_photos: List[Path] = []
    if photo_dir.exists() and photo_dir.is_dir():
        all_photos = sorted(
            [
                p
                for p in photo_dir.rglob("*")
                if p.is_file() and p.suffix.lower() in IMAGE_EXTS and not p.name.startswith(".")
            ]
        )

    assigned: Dict[str, Optional[Path]] = {slot[0]: None for slot in PHOTO_SLOTS}
    used: set[Path] = set()

    # 第一轮：按文件名关键词匹配
    for key, _, _, keywords, _ in PHOTO_SLOTS:
        for photo in all_photos:
            if photo in used:
                continue
            stem_lower = photo.stem.lower()
            if any(kw.lower() in stem_lower for kw in keywords):
                assigned[key] = photo
                used.add(photo)
                break

    # 第二轮：用未分配的照片按顺序填补空缺
    remaining = [p for p in all_photos if p not in used]
    rem_idx = 0
    for key, _, _, _, _ in PHOTO_SLOTS:
        if assigned[key] is None and rem_idx < len(remaining):
            assigned[key] = remaining[rem_idx]
            rem_idx += 1

    # 第三轮：若照片总数少于槽位数但至少有1张照片，仅在不希望留白时可留作占位或复用
    return assigned, all_photos


def render_photo_box(
    photo_path: Optional[Path],
    target_w_in: float,
    target_h_in: float,
    title: str,
    subtitle: str,
    base_rgb: Tuple[int, int, int],
    slot_key: str,
) -> io.BytesIO:
    """将真实照片按目标宽高比居中裁剪（避免拉伸变形），或生成精美示意图。"""
    dpi = 220
    px_w = max(400, int(target_w_in * dpi))
    px_h = max(300, int(target_h_in * dpi))

    if photo_path is not None and photo_path.exists():
        try:
            with Image.open(photo_path) as im:
                im = ImageOps.exif_transpose(im).convert("RGB")
                fitted = ImageOps.fit(im, (px_w, px_h), method=Image.Resampling.LANCZOS)
                buf = io.BytesIO()
                fitted.save(buf, format="JPEG", quality=92)
                buf.seek(0)
                return buf
        except Exception as exc:
            print(f"[提示] 读取照片 {photo_path} 失败 ({exc})，使用精美占位图代替。")

    # 生成高质感占位示意图（当在云端或未挂载本地照片目录时显示）
    img = Image.new("RGB", (px_w, px_h), base_rgb)
    draw = ImageDraw.Draw(img)

    r0, g0, b0 = base_rgb
    for y in range(px_h):
        ratio = y / max(1, px_h - 1)
        r = min(255, int(r0 * (1 - 0.28 * ratio) + 25 * ratio))
        g = min(255, int(g0 * (1 - 0.22 * ratio) + 40 * ratio))
        b = min(255, int(b0 * (1 - 0.18 * ratio) + 55 * ratio))
        draw.line([(0, y), (px_w, y)], fill=(r, g, b))

    # 装饰性几何背景与相机取景框
    pad = int(min(px_w, px_h) * 0.06)
    draw.rounded_rectangle(
        [pad, pad, px_w - pad, px_h - pad],
        radius=18,
        outline=(255, 255, 255),
        width=2,
    )

    # 绘制简约取景图标
    cx, cy = px_w // 2, int(px_h * 0.34)
    cam_w, cam_h = int(min(px_w, px_h) * 0.22), int(min(px_w, px_h) * 0.16)
    draw.rounded_rectangle(
        [cx - cam_w // 2, cy - cam_h // 2, cx + cam_w // 2, cy + cam_h // 2],
        radius=10,
        outline=(255, 255, 255),
        width=3,
    )
    lens_r = int(min(cam_w, cam_h) * 0.28)
    draw.ellipse([cx - lens_r, cy - lens_r, cx + lens_r, cy + lens_r], outline=(255, 255, 255), width=3)

    font_title = _load_cjk_font(max(22, int(min(px_w, px_h) * 0.065)))
    font_sub = _load_cjk_font(max(15, int(min(px_w, px_h) * 0.042)))
    font_hint = _load_cjk_font(max(13, int(min(px_w, px_h) * 0.035)))

    draw.text((cx, int(px_h * 0.54)), f"【照片位】{title}", font=font_title, fill=(255, 255, 255), anchor="mm")
    draw.text((cx, int(px_h * 0.65)), subtitle, font=font_sub, fill=(232, 240, 248), anchor="mm")
    draw.text(
        (cx, int(px_h * 0.78)),
        f"运行脚本自动载入 5周年旅游照片（推荐命名含：{slot_key} 或景点名）",
        font=font_hint,
        fill=(210, 225, 240),
        anchor="mm",
    )

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=90)
    buf.seek(0)
    return buf


def _set_run_font(run, text: str, size: float, color_hex: str, bold: bool = False):
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color_hex)
    run.font.name = FONT_FAMILY
    rpr = run._r.get_or_add_rPr()
    rpr.set("lang", "zh-CN")
    rpr.set("altLang", "en-US")
    for tag in ("latin", "ea", "cs"):
        el = rpr.find(qn(f"a:{tag}"))
        if el is None:
            el = rpr.makeelement(qn(f"a:{tag}"), {"typeface": FONT_FAMILY})
            rpr.append(el)
        else:
            el.set("typeface", FONT_FAMILY)


def add_textbox(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    paragraphs: List[dict],
    valign: str = "top",
):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    body = tf._txBody.find(qn("a:bodyPr"))
    body.set("anchor", {"top": "t", "middle": "ctr", "bottom": "b"}[valign])
    for key in ("lIns", "rIns", "tIns", "bIns"):
        body.set(key, "0")

    align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    for idx, p_info in enumerate(paragraphs):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align_map.get(p_info.get("align", "left"), PP_ALIGN.LEFT)
        if "line_spacing" in p_info:
            p.line_spacing = p_info["line_spacing"]
        if "space_before" in p_info:
            p.space_before = Pt(p_info["space_before"])
        if "space_after" in p_info:
            p.space_after = Pt(p_info["space_after"])

        if "runs" in p_info:
            for r_info in p_info["runs"]:
                run = p.add_run()
                _set_run_font(
                    run,
                    r_info["text"],
                    r_info.get("size", 14),
                    r_info.get("color", INK),
                    r_info.get("bold", False),
                )
        else:
            run = p.add_run()
            _set_run_font(
                run,
                p_info["text"],
                p_info.get("size", 14),
                p_info.get("color", INK),
                p_info.get("bold", False),
            )
    return shape


def add_card(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    fill_hex: str,
    border_hex: Optional[str] = None,
    rounded: bool = True,
):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded and hasattr(shape, "adjustments") and len(shape.adjustments) > 0:
        shape.adjustments[0] = 0.03
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill_hex)
    if border_hex:
        shape.line.color.rgb = rgb(border_hex)
        shape.line.width = Pt(1.0)
    else:
        shape.line.fill.background()
    return shape


def set_slide_bg(slide, color_hex: str):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color_hex)


def add_header(slide, badge_text: str, title_text: str, subtitle_text: str, page_num: int, total_pages: int = 7):
    # 顶部品牌色细条
    add_card(slide, 0, 0, SLIDE_W, 0.08, ATOMY_BLUE, rounded=False)

    # 胶囊标签
    add_card(slide, 0.65, 0.35, 2.2, 0.36, ATOMY_SOFT, rounded=True)
    add_textbox(
        slide,
        0.65,
        0.35,
        2.2,
        0.36,
        [{"text": badge_text, "size": 11, "color": ATOMY_BLUE, "bold": True, "align": "center"}],
        valign="middle",
    )

    # 标题与副标题
    add_textbox(
        slide,
        0.65,
        0.78,
        9.8,
        0.52,
        [{"text": title_text, "size": 24, "color": INK, "bold": True}],
    )
    add_textbox(
        slide,
        0.65,
        1.28,
        10.5,
        0.34,
        [{"text": subtitle_text, "size": 13, "color": MUTED}],
    )

    # 右上角品牌标识与底部页码
    add_textbox(
        slide,
        10.2,
        0.40,
        2.5,
        0.32,
        [{"text": "ATOMY · 艾多美5周年员工旅游", "size": 11, "color": MUTED, "align": "right", "bold": True}],
    )
    add_textbox(
        slide,
        11.2,
        7.05,
        1.5,
        0.28,
        [{"text": f"{page_num:02d} / {total_pages:02d}", "size": 10.5, "color": MUTED, "align": "right"}],
    )


def add_photo_card(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    slot_key: str,
    assigned: Dict[str, Optional[Path]],
    slot_meta: Dict[str, Tuple[str, str, List[str], Tuple[int, int, int]]],
    caption_title: Optional[str] = None,
    caption_desc: Optional[str] = None,
):
    """添加带白色相框底衬与可选底部文字说明的照片卡片。"""
    add_card(slide, x, y, w, h, WHITE, BORDER, rounded=True)
    pad = 0.10
    caption_h = 0.72 if (caption_title or caption_desc) else 0.0
    img_w = w - pad * 2
    img_h = h - pad * 2 - caption_h

    title, subtitle, _, base_rgb = slot_meta[slot_key]
    img_stream = render_photo_box(
        assigned.get(slot_key),
        img_w,
        img_h,
        title,
        subtitle,
        base_rgb,
        slot_key,
    )
    slide.shapes.add_picture(img_stream, Inches(x + pad), Inches(y + pad), Inches(img_w), Inches(img_h))

    if caption_title or caption_desc:
        paras = []
        if caption_title:
            paras.append({"text": caption_title, "size": 13, "color": INK, "bold": True, "space_after": 2})
        if caption_desc:
            paras.append({"text": caption_desc, "size": 11, "color": BODY})
        add_textbox(slide, x + pad + 0.05, y + h - caption_h + 0.04, img_w - 0.1, caption_h - 0.08, paras, valign="middle")


def build_presentation(photo_dir: Path, output_path: Path) -> Tuple[int, List[Path]]:
    assigned, all_photos = discover_and_assign_photos(photo_dir)
    slot_meta = {s[0]: (s[1], s[2], s[3], s[4]) for s in PHOTO_SLOTS}

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    blank = prs.slide_layouts[6]
    total_pages = 7

    # =========================================================================
    # 第 1 页：封面页 —— 感恩艾多美 · 五载同行 · 沪杭幸福之旅
    # =========================================================================
    s1 = prs.slides.add_slide(blank)
    set_slide_bg(s1, BG_DARK)
    add_card(s1, 0, 0, SLIDE_W, 0.10, GOLD, rounded=False)

    # 左侧标题与引言
    add_card(s1, 0.75, 0.85, 3.1, 0.42, ATOMY_BLUE, rounded=True)
    add_textbox(
        s1,
        0.75,
        0.85,
        3.1,
        0.42,
        [{"text": "ATOMY · 入职满5周年福利旅游", "size": 12, "color": WHITE, "bold": True, "align": "center"}],
        valign="middle",
    )

    add_textbox(
        s1,
        0.75,
        1.55,
        6.2,
        1.7,
        [
            {"text": "感恩艾多美 · 五载同行", "size": 36, "color": WHITE, "bold": True, "space_after": 8},
            {"text": "携家人共赴上海·杭州幸福之旅", "size": 24, "color": GOLD, "bold": True},
        ],
    )

    add_textbox(
        s1,
        0.75,
        3.25,
        6.0,
        1.25,
        [
            {
                "text": "在艾多美，每满 3、5、7、10 年即可享受公司承担费用的专属员工旅游。\n今年迎来了我入职满 5 周年的重要里程碑，带着父母、老婆和孩子一同走进繁华上海与诗画杭州，把公司的关怀化作全家人最温暖的记忆。",
                "size": 14,
                "color": "DCE7F3",
                "line_spacing": 1.45,
            }
        ],
    )

    # 底部3个核心信息胶囊卡
    cover_pills = [
        ("福利机制", "满 3 / 5 / 7 / 10 年\n费用由公司全额承担"),
        ("同行家人", "父母 · 老婆 · 孩子\n三代同堂温馨出游"),
        ("双城足迹", "上海外滩 · 金茂56层 · 豫园\n杭州西湖 · 雷峰塔 · 海鲜盛宴"),
    ]
    for idx, (p_title, p_desc) in enumerate(cover_pills):
        px = 0.75 + idx * 2.05
        add_card(s1, px, 4.95, 1.92, 1.45, "173C63", "2A5685", rounded=True)
        add_textbox(
            s1,
            px + 0.14,
            5.10,
            1.64,
            1.15,
            [
                {"text": p_title, "size": 12.5, "color": GOLD, "bold": True, "space_after": 6},
                {"text": p_desc, "size": 11, "color": WHITE, "line_spacing": 1.35},
            ],
        )

    # 右侧大图：全家福
    add_photo_card(
        s1,
        7.15,
        0.85,
        5.45,
        5.65,
        "cover_family",
        assigned,
        slot_meta,
        caption_title="五周年全家福 · 最珍贵的礼物是家人的笑容",
        caption_desc="感谢艾多美承担旅途费用，让我们一家老小尽享天伦之乐",
    )
    add_textbox(
        s1,
        11.2,
        7.05,
        1.5,
        0.28,
        [{"text": f"01 / {total_pages:02d}", "size": 10.5, "color": "8FA8C4", "align": "right"}],
    )

    # =========================================================================
    # 第 2 页：公司暖心福利与行程总览 —— 3/5/7/10年里程碑与三代同堂路线
    # =========================================================================
    s2 = prs.slides.add_slide(blank)
    set_slide_bg(s2, BG_LIGHT)
    add_header(
        s2,
        "暖心福利 · 行程概览",
        "艾多美周年旅游制度：见证成长，更犒赏身后的家人",
        "每满 3、5、7、10 年享公司全额承担旅游福利 —— 我的第 5 年，圆了一场三代同堂的江南行",
        2,
        total_pages,
    )

    # 上半部分：3 / 5 / 7 / 10 年阶梯里程碑
    milestones = [
        ("满 3 周年", "初心扎根 · 启程", "融入艾多美大家庭\n开启第一次周年福利之旅", False),
        ("满 5 周年（本次）", "五载深耕 · 阖家同游", "带父母、老婆、孩子畅游沪杭\n旅游费用由公司暖心承担", True),
        ("满 7 周年", "笃行致远 · 进阶", "七年相伴与公司共同成长\n期待下一场远方风景", False),
        ("满 10 周年", "十年同心 · 荣耀", "十年里程碑盛大纪念\n感恩长期主义的双向奔赴", False),
    ]
    for idx, (m_year, m_tag, m_desc, is_active) in enumerate(milestones):
        mx = 0.65 + idx * 3.05
        fill_c = ATOMY_BLUE if is_active else WHITE
        border_c = GOLD if is_active else BORDER
        title_c = WHITE if is_active else ATOMY_DARK
        tag_c = GOLD_SOFT if is_active else ATOMY_BLUE
        desc_c = WHITE if is_active else BODY

        add_card(s2, mx, 1.80, 2.88, 1.75, fill_c, border_c, rounded=True)
        add_textbox(
            s2,
            mx + 0.20,
            1.95,
            2.48,
            1.45,
            [
                {"text": m_year, "size": 16, "color": title_c, "bold": True, "space_after": 4},
                {"text": m_tag, "size": 12.5, "color": tag_c, "bold": True, "space_after": 6},
                {"text": m_desc, "size": 11, "color": desc_c, "line_spacing": 1.35},
            ],
        )

    # 下半部分左卡片：同行家人与初心
    add_card(s2, 0.65, 3.82, 5.85, 3.0, WHITE, BORDER, rounded=True)
    add_card(s2, 0.65, 3.82, 0.10, 3.0, GOLD, rounded=False)
    add_textbox(
        s2,
        0.95,
        4.02,
        5.35,
        2.6,
        [
            {"text": "一家五口的幸福同行", "size": 18, "color": INK, "bold": True, "space_after": 10},
            {
                "text": "• 同行成员：父母、老婆、孩子与我（三代同堂）\n• 费用保障：由艾多美公司承担，出行轻松无负担\n• 旅途初心：平时忙于工作，难得陪父母和妻儿远行。趁着 5 周年员工旅游福利，精心规划了「上海 + 杭州」双城慢节奏路线，兼顾老人的舒适度与孩子的探索欲，让全家人共享高品质的团聚时光。",
                "size": 13,
                "color": BODY,
                "line_spacing": 1.48,
            },
        ],
    )

    # 下半部分右卡片：双城精华打卡清单
    add_card(s2, 6.80, 3.82, 5.88, 3.0, WHITE, BORDER, rounded=True)
    add_card(s2, 6.80, 3.82, 0.10, 3.0, ATOMY_BLUE, rounded=False)
    add_textbox(
        s2,
        7.10,
        4.02,
        5.35,
        2.6,
        [
            {"text": "上海 · 杭州双城精华行程", "size": 18, "color": INK, "bold": True, "space_after": 10},
            {
                "runs": [
                    {"text": "第一站 · 魔都上海（都市地标与海派风情）\n", "size": 13.5, "color": ATOMY_BLUE, "bold": True},
                    {
                        "text": "外滩万国建筑群 → 金茂大厦56层观景自助（300元/位） → 豫园江南园林 → 城隍庙民俗小吃 → 南京路步行街\n\n",
                        "size": 12.5,
                        "color": BODY,
                    },
                    {"text": "第二站 · 诗画杭州（山水人文与海鲜盛宴）\n", "size": 13.5, "color": GOLD, "bold": True},
                    {
                        "text": "西湖漫步赏景 → 雷峰塔登高望远 → 丰盛海鲜大餐（大龙虾、膏满面包蟹等）",
                        "size": 12.5,
                        "color": BODY,
                    },
                ],
                "line_spacing": 1.38,
            },
        ],
    )

    # =========================================================================
    # 第 3 页：上海站（一） —— 外滩江风与金茂大厦56层300元/位云端观景自助
    # =========================================================================
    s3 = prs.slides.add_slide(blank)
    set_slide_bg(s3, BG_LIGHT)
    add_header(
        s3,
        "上海站 · 璀璨天际",
        "漫步外滩江畔，登顶金茂大厦56层享云端观景自助",
        "吹拂黄浦江畔晚风，在 56 层高空一边俯瞰魔都繁华全景，一边与家人举杯共享 300 元/位星级自助美食",
        3,
        total_pages,
    )

    # 左上：外滩大图
    add_photo_card(
        s3,
        0.65,
        1.78,
        5.90,
        5.05,
        "bund",
        assigned,
        slot_meta,
        caption_title="上海外滩 · 浦江两岸风云际会",
        caption_desc="带父母和孩子在外滩合影留念，看万国建筑与陆家嘴高楼交相辉映",
    )

    # 右上：金茂大厦56层照片
    add_photo_card(
        s3,
        6.80,
        1.78,
        5.88,
        3.25,
        "jinmao",
        assigned,
        slot_meta,
        caption_title="金茂大厦 56 层 · 300元/位高空观景自助餐",
        caption_desc="穿行于壮观的空中中庭，在云端视野下享用琳琅满目的环球美食",
    )

    # 右下：体验心得卡片
    add_card(s3, 6.80, 5.20, 5.88, 1.63, WHITE, BORDER, rounded=True)
    add_card(s3, 6.80, 5.20, 5.88, 0.06, GOLD, rounded=False)
    add_textbox(
        s3,
        7.05,
        5.35,
        5.40,
        1.35,
        [
            {"text": "旅途高光时刻 · 父母与孩子的惊喜赞叹", "size": 14, "color": ATOMY_DARK, "bold": True, "space_after": 5},
            {
                "text": "站在金茂大厦 56 层落地窗前俯瞰外滩与黄浦江，300 元/位的高品质观景自助餐菜品丰富精致。看着父母开心地品尝美食、孩子兴奋地指着窗外的东方明珠，由衷感谢公司提供的优厚旅游支持！",
                "size": 12,
                "color": BODY,
                "line_spacing": 1.38,
            },
        ],
    )

    # =========================================================================
    # 第 4 页：上海站（二） —— 豫园、城隍庙与南京路的海派烟火气
    # =========================================================================
    s4 = prs.slides.add_slide(blank)
    set_slide_bg(s4, BG_LIGHT)
    add_header(
        s4,
        "上海站 · 海派风情",
        "豫园寻幽、城隍庙觅食、南京路漫步：品味老上海与新繁华",
        "从江南古典园林的亭台楼阁，到十里洋场南京路的热闹街景，一步一景皆是团圆欢笑",
        4,
        total_pages,
    )

    sh_spots = [
        (
            "yuyuan",
            "豫园 · 江南古典园林雅韵",
            "曲径通幽、奇石亭台，父母最喜欢这份古色古香的江南韵味，慢步其间格外舒心。",
        ),
        (
            "chenghuangmiao",
            "城隍庙 · 地道海派民俗与小吃",
            "九曲桥畔热闹非凡，一家人品尝南翔小笼与特色糕点，感受浓郁的市井烟火气。",
        ),
        (
            "nanjinglu",
            "南京路步行街 · 十里洋场百年繁华",
            "老字号百货与现代潮流交织，陪老婆孩子逛街打卡、选购伴手礼，满载欢声笑语。",
        ),
    ]

    for idx, (s_key, s_title, s_desc) in enumerate(sh_spots):
        sx = 0.65 + idx * 4.10
        add_photo_card(
            s4,
            sx,
            1.78,
            3.83,
            3.70,
            s_key,
            assigned,
            slot_meta,
        )
        add_card(s4, sx, 5.62, 3.83, 1.22, WHITE, BORDER, rounded=True)
        add_textbox(
            s4,
            sx + 0.18,
            5.74,
            3.47,
            1.0,
            [
                {"text": s_title, "size": 13.5, "color": ATOMY_DARK, "bold": True, "space_after": 4},
                {"text": s_desc, "size": 11.5, "color": BODY, "line_spacing": 1.32},
            ],
        )

    # =========================================================================
    # 第 5 页：杭州站 —— 西湖泛舟漫步与雷峰塔登高揽胜
    # =========================================================================
    s5 = prs.slides.add_slide(blank)
    set_slide_bg(s5, BG_LIGHT)
    add_header(
        s5,
        "杭州站 · 诗画江南",
        "西湖烟雨入画来，登临雷峰塔俯瞰一城山色半城湖",
        "告别魔都的摩登节奏，带着家人来到温婉灵秀的杭州，在湖光山色间放松身心",
        5,
        total_pages,
    )

    add_photo_card(
        s5,
        0.65,
        1.78,
        5.85,
        3.95,
        "xihu",
        assigned,
        slot_meta,
        caption_title="杭州西湖 · 淡妆浓抹总相宜",
        caption_desc="全家人沿着西湖畔悠闲漫步，杨柳依依、碧波荡漾，令人心旷神怡",
    )

    add_photo_card(
        s5,
        6.82,
        1.78,
        5.86,
        3.95,
        "leifengta",
        assigned,
        slot_meta,
        caption_title="雷峰塔 · 登高极目揽西湖全景",
        caption_desc="乘扶梯轻松登塔，对老人和孩子非常友好；凭栏远眺，西湖美景尽收眼底",
    )

    # 底部通栏心得条
    add_card(s5, 0.65, 5.90, 12.03, 0.95, ATOMY_SOFT, BORDER, rounded=True)
    add_textbox(
        s5,
        0.90,
        6.02,
        11.5,
        0.72,
        [
            {
                "runs": [
                    {"text": "杭州站随感：", "size": 13, "color": ATOMY_BLUE, "bold": True},
                    {
                        "text": "以前只在课本和电视里看到的西湖与雷峰塔，这次终于带着父母、老婆和孩子实地打卡。慢节奏的山水之旅洗去了平日的疲惫，也让孩子在旅途中收获了生动的人文历史课堂。",
                        "size": 12.5,
                        "color": INK,
                    },
                ],
                "line_spacing": 1.35,
            }
        ],
        valign="middle",
    )

    # =========================================================================
    # 第 6 页：舌尖上的幸福 —— 龙虾、面包蟹等海鲜大餐犒赏家人
    # =========================================================================
    s6 = prs.slides.add_slide(blank)
    set_slide_bg(s6, BG_LIGHT)
    add_header(
        s6,
        "美食记忆 · 舌尖盛宴",
        "海鲜大餐狂欢：鲜美大龙虾与膏满面包蟹，共享团圆滋味",
        "旅途中不仅有眼前的风景，更有舌尖上的幸福 —— 用丰盛的海鲜盛宴好好犒劳辛苦付出的家人",
        6,
        total_pages,
    )

    add_photo_card(
        s6,
        0.65,
        1.78,
        4.35,
        5.05,
        "lobster",
        assigned,
        slot_meta,
        caption_title="招牌大龙虾 · 肉质Q弹鲜甜",
        caption_desc="个头十足的大龙虾端上桌，仪式感拉满，全家人大快朵颐",
    )

    add_photo_card(
        s6,
        5.20,
        1.78,
        4.35,
        5.05,
        "crab",
        assigned,
        slot_meta,
        caption_title="黄金面包蟹 · 蟹膏饱满馥郁",
        caption_desc="满满的蟹黄与鲜嫩蟹肉，老人和孩子都吃得津津有味",
    )

    # 右侧美食清单与幸福感言卡片
    add_card(s6, 9.75, 1.78, 2.93, 5.05, WHITE, BORDER, rounded=True)
    add_card(s6, 9.75, 1.78, 2.93, 0.08, GOLD, rounded=False)
    add_textbox(
        s6,
        9.98,
        2.02,
        2.48,
        4.6,
        [
            {"text": "舌尖高光菜单", "size": 17, "color": INK, "bold": True, "space_after": 12},
            {"text": "• 金茂56层 300元/位观景自助", "size": 12.5, "color": ATOMY_BLUE, "bold": True, "space_after": 6},
            {"text": "高空江景佐餐，中西料理、甜品海鲜一应俱全\n", "size": 11.5, "color": BODY, "space_after": 10},
            {"text": "• 鲜活大龙虾盛宴", "size": 12.5, "color": GOLD, "bold": True, "space_after": 6},
            {"text": "鲜嫩多汁、营养丰富，全家人的最爱\n", "size": 11.5, "color": BODY, "space_after": 10},
            {"text": "• 爆膏面包蟹 & 时令海鲜", "size": 12.5, "color": GOLD, "bold": True, "space_after": 6},
            {"text": "蟹黄醇厚、鲜味十足，孩子连连称赞\n", "size": 11.5, "color": BODY, "space_after": 14},
            {
                "text": "“最好的美食，是和最爱的人一起分享；最暖的福利，是公司把关怀延伸到了每个小家庭。”",
                "size": 11.5,
                "color": ATOMY_DARK,
                "bold": True,
                "line_spacing": 1.38,
            },
        ],
    )

    # =========================================================================
    # 第 7 页：旅游心得与感恩展望 —— 感恩艾多美，奔赴下一个7年、10年
    # =========================================================================
    s7 = prs.slides.add_slide(blank)
    set_slide_bg(s7, BG_LIGHT)
    add_header(
        s7,
        "旅游心得 · 感恩致远",
        "五年成长心怀感恩，带着满满幸福能量再出发",
        "一次有温度的员工旅游，连起了公司与家庭的纽带，也坚定了未来继续并肩前行的步伐",
        7,
        total_pages,
    )

    reflections = [
        (
            "01 · 感恩公司福利，让奋斗更有温度",
            ATOMY_BLUE,
            "艾多美每满 3、5、7、10 年由公司承担费用的员工旅游制度，不仅是对员工五年辛勤付出的认可，更是对员工家属最实在的馈赠。看到父母和妻儿在旅途中的灿烂笑容，我作为艾多美一员倍感自豪。",
        ),
        (
            "02 · 珍视家人陪伴，凝聚小家幸福力量",
            GOLD,
            "从上海外滩、金茂大厦 56 层观景自助、豫园城隍庙南京路，到杭州西湖、雷峰塔与丰盛的海鲜大餐，这趟旅程弥补了平日忙碌的陪伴时光。家人的理解与支持，始终是我安心工作的最大底气。",
        ),
        (
            "03 · 蓄力再出发，相约第 7 年与第 10 年",
            ATOMY_DARK,
            "五周年既是里程碑，更是新起点。回到工作岗位后，我将把旅途中收获的感动与活力转化为工作热忱，脚踏实地创造更大价值，期待满 7 年、满 10 年时，再次带着家人开启新的美好旅程！",
        ),
    ]

    for idx, (r_title, r_color, r_body) in enumerate(reflections):
        rx = 0.65 + idx * 4.10
        add_card(s7, rx, 1.82, 3.83, 3.35, WHITE, BORDER, rounded=True)
        add_card(s7, rx, 1.82, 3.83, 0.08, r_color, rounded=False)
        add_textbox(
            s7,
            rx + 0.24,
            2.08,
            3.35,
            2.85,
            [
                {"text": r_title, "size": 15, "color": r_color, "bold": True, "space_after": 12},
                {"text": r_body, "size": 12.5, "color": BODY, "line_spacing": 1.48},
            ],
        )

    # 底部结语横幅
    add_card(s7, 0.65, 5.45, 12.03, 1.35, BG_DARK, rounded=True)
    add_textbox(
        s7,
        0.95,
        5.60,
        11.4,
        1.05,
        [
            {
                "text": "心怀感恩 · 与艾多美共赴下一个五年、十年！",
                "size": 20,
                "color": WHITE,
                "bold": True,
                "align": "center",
                "space_after": 6,
            },
            {
                "text": "感谢公司领导与同事们的关怀与支持 · 5周年员工旅游心得分享",
                "size": 12.5,
                "color": GOLD,
                "align": "center",
            },
        ],
        valign="middle",
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return len(prs.slides), all_photos


def main():
    parser = argparse.ArgumentParser(description="生成艾多美5周年员工旅游心得PPT（自动配图）")
    parser.add_argument(
        "--photos",
        type=str,
        default=DEFAULT_PHOTO_DIR,
        help=f"照片文件夹路径（默认：{DEFAULT_PHOTO_DIR}）",
    )
    parser.add_argument(
        "--output",
        type=str,
        default=DEFAULT_OUTPUT,
        help=f"输出PPT文件路径（默认：{DEFAULT_OUTPUT}）",
    )
    args = parser.parse_args()

    photo_dir = Path(args.photos).expanduser()
    output_path = Path(args.output).expanduser()

    slide_count, found_photos = build_presentation(photo_dir, output_path)
    print(f"[完成] 已生成 {slide_count} 页 PPT：{output_path.resolve()}")
    if found_photos:
        print(f"[配图] 从目录「{photo_dir}」成功识别并嵌入 {len(found_photos)} 张照片。")
    else:
        print(
            f"[提示] 当前环境未检测到目录「{photo_dir}」中的照片，已生成带精美取景框的图文版 PPT。\n"
            f"       在本机 Mac 上运行 `python3 generate_travel_ppt.py` 即可一键将真实照片自动裁切嵌入！"
        )


if __name__ == "__main__":
    main()
