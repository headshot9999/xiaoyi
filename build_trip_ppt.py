#!/usr/bin/env python3
"""
艾多美5周年员工旅游心得 PPT 生成器（上海 · 杭州）

支持自动扫描本地照片文件夹（默认优先检查：
1. 命令行参数 --photos 指定的目录
2. /Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片
3. 当前目录下的 ./5周年旅游照片 或 ./photos
）
并将照片按文件名关键字或顺序智能居中裁剪后排入对应页面；
若在云端环境或未找到照片时，会自动生成精美的照片占位框，
您也可以直接在 PowerPoint 中右键图片选择「更改图片」一键替换，
或在本地电脑运行本脚本自动嵌入全部照片！
"""

from __future__ import annotations

import argparse
import io
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# 字体配置（兼容 macOS 苹方/思源黑体 与 Windows 微软雅黑）
FONT_SANS = "PingFang SC"
FONT_BOLD = "PingFang SC Semibold"

# 艾多美清新蓝 + 江南温润暖金配色方案
WHITE = "FFFFFF"
BG_SOFT = "F5F8FB"       # 柔和浅云蓝背景
BG_CARD = "FFFFFF"       # 卡片白
ATOMY_BLUE = "0093D0"    # 艾多美品牌蓝
ATOMY_DARK = "0C2340"    # 深海藏青（封面与强调）
ATOMY_LIGHT = "E6F4FA"   # 浅蓝高亮底
GOLD = "C8923A"          # 暖金（5周年/家庭温馨点缀）
GOLD_SOFT = "FDF6EB"     # 暖米金底
INK = "1A2433"           # 主标题深色
BODY = "3D4B5C"          # 正文灰蓝
MUTED = "7A8B9E"         # 辅助说明文字
BORDER = "DCE5EF"        # 细边框色
PHOTO_BG = "EBF1F6"      # 照片占位底色

W = 13.333333
H = 7.5
ML = 0.68

DEFAULT_MAC_PHOTO_DIR = Path("/Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片")
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff"}


def rgb(hex_str: str) -> RGBColor:
    return RGBColor.from_string(hex_str)


def _anchor(tf, valign: str = "top") -> None:
    body = tf._txBody.find(qn("a:bodyPr"))
    body.set("anchor", {"top": "t", "middle": "ctr", "bottom": "b"}[valign])
    for key in ("lIns", "rIns", "tIns", "bIns"):
        body.set(key, "0")


def _run(p, text: str, font: str, size: float, color: str, bold: bool = False) -> None:
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


def tb(slide, x: float, y: float, w: float, h: float, paragraphs: list[dict], valign: str = "top"):
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
                    item.get("font", FONT_SANS),
                    item["size"],
                    item["color"],
                    item.get("bold", False),
                )
        else:
            _run(
                p,
                info["text"],
                info.get("font", FONT_SANS),
                info["size"],
                info["color"],
                info.get("bold", False),
            )
    return shape


def rect(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    fill: str,
    line_color: str | None = None,
    line_width: float = 0.75,
    rounded: bool = False,
):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded and hasattr(shape, "adjustments") and len(shape.adjustments) > 0:
        shape.adjustments[0] = 0.03
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    if line_color:
        shape.line.color.rgb = rgb(line_color)
        shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape


def bg(slide, color: str) -> None:
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color)


def notes(slide, text: str) -> None:
    slide.notes_slide.notes_text_frame.text = text


def new_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def header(slide, page_num: int, total_pages: int, tag: str, title: str, subtitle: str = "") -> None:
    # 顶部细装饰条
    rect(slide, 0, 0, W, 0.06, ATOMY_BLUE)
    # 章节胶囊标签
    rect(slide, ML, 0.34, 2.25, 0.32, ATOMY_LIGHT, rounded=True)
    tb(
        slide,
        ML,
        0.34,
        2.25,
        0.32,
        [{"text": tag, "size": 11.5, "color": ATOMY_BLUE, "bold": True, "align": "center"}],
        valign="middle",
    )
    if subtitle:
        tb(
            slide,
            ML + 2.40,
            0.34,
            8.2,
            0.32,
            [{"text": subtitle, "size": 11.5, "color": MUTED}],
            valign="middle",
        )
    # 页码与品牌角标
    tb(
        slide,
        W - ML - 3.2,
        0.34,
        3.2,
        0.32,
        [
            {
                "text": f"ATOMY 5周年员工旅游心得  ·  {page_num:02d} / {total_pages:02d}",
                "size": 10.5,
                "color": MUTED,
                "align": "right",
            }
        ],
        valign="middle",
    )
    # 主标题
    tb(
        slide,
        ML,
        0.74,
        W - ML * 2,
        0.52,
        [{"text": title, "size": 25, "color": INK, "bold": True}],
        valign="middle",
    )
    rect(slide, ML, 1.32, W - ML * 2, 0.015, BORDER)


def crop_image_to_stream(img_path: Path, target_w_in: float, target_h_in: float, dpi: int = 220) -> io.BytesIO:
    """读取照片、修正 EXIF 方向，并按目标宽高比居中裁剪，返回内存字节流。"""
    with Image.open(img_path) as im:
        im = ImageOps.exif_transpose(im)
        if im.mode not in ("RGB", "RGBA"):
            im = im.convert("RGB")
        target_ratio = target_w_in / target_h_in
        w, h = im.size
        src_ratio = w / h
        if src_ratio > target_ratio:
            new_w = int(h * target_ratio)
            left = (w - new_w) // 2
            im = im.crop((left, 0, left + new_w, h))
        else:
            new_h = int(w / target_ratio)
            top = (h - new_h) // 2
            im = im.crop((0, top, w, top + new_h))

        out_w = max(400, int(target_w_in * dpi))
        out_h = max(300, int(target_h_in * dpi))
        if im.size[0] > out_w:
            im = im.resize((out_w, out_h), Image.Resampling.LANCZOS)

        buf = io.BytesIO()
        if im.mode == "RGBA":
            im = im.convert("RGB")
        im.save(buf, format="JPEG", quality=90)
        buf.seek(0)
        return buf


class PhotoAssigner:
    """从照片目录中按关键词或顺序分配照片。"""

    def __init__(self, photo_dir: Path | None):
        self.photo_dir = photo_dir
        self.all_photos: list[Path] = []
        self.used: set[Path] = set()
        if photo_dir and photo_dir.exists() and photo_dir.is_dir():
            self.all_photos = sorted(
                [
                    p
                    for p in photo_dir.rglob("*")
                    if p.is_file() and p.suffix.lower() in IMAGE_EXTS and not p.name.startswith(".")
                ]
            )

    def pick(self, keywords: Iterable[str] = ()) -> Path | None:
        if not self.all_photos:
            return None
        # 1. 优先匹配文件名包含关键字且未使用的照片
        for kw in keywords:
            for p in self.all_photos:
                if p not in self.used and kw.lower() in p.stem.lower():
                    self.used.add(p)
                    return p
        # 2. 按顺序取尚未使用的照片
        for p in self.all_photos:
            if p not in self.used:
                self.used.add(p)
                return p
        # 3. 如照片总数少于槽位，循环复用
        idx = len(self.used) % len(self.all_photos)
        return self.all_photos[idx]


def photo_slot(
    slide,
    assigner: PhotoAssigner,
    x: float,
    y: float,
    w: float,
    h: float,
    caption: str,
    hint: str,
    keywords: Iterable[str] = (),
    dark_mode: bool = False,
) -> None:
    """绘制带相框与图注的照片区域；若找到实拍照片则自动裁剪嵌入，否则生成优雅占位框。"""
    # 外层白底相框卡片
    frame_bg = "162F4D" if dark_mode else WHITE
    frame_border = "2C4C70" if dark_mode else BORDER
    caption_color = "D6E6F5" if dark_mode else INK
    sub_color = "8FAEC9" if dark_mode else MUTED

    rect(slide, x, y, w, h, frame_bg, frame_border, line_width=1.0, rounded=True)

    pad = 0.10
    cap_h = 0.42
    img_x = x + pad
    img_y = y + pad
    img_w = w - pad * 2
    img_h = h - pad * 2 - cap_h

    matched_photo = assigner.pick(keywords)
    if matched_photo is not None:
        try:
            stream = crop_image_to_stream(matched_photo, img_w, img_h)
            slide.shapes.add_picture(stream, Inches(img_x), Inches(img_y), Inches(img_w), Inches(img_h))
        except Exception:
            matched_photo = None

    if matched_photo is None:
        # 精美占位区
        slot_fill = "1E3B5E" if dark_mode else PHOTO_BG
        slot_line = "3A628F" if dark_mode else "BFD2E3"
        rect(slide, img_x, img_y, img_w, img_h, slot_fill, slot_line, line_width=0.85)
        # 顶部装饰角标
        rect(slide, img_x + 0.14, img_y + 0.14, 0.95, 0.24, ATOMY_BLUE, rounded=True)
        tb(
            slide,
            img_x + 0.14,
            img_y + 0.14,
            0.95,
            0.24,
            [{"text": "实拍照片位", "size": 9.5, "color": WHITE, "bold": True, "align": "center"}],
            valign="middle",
        )
        tb(
            slide,
            img_x + 0.2,
            img_y + 0.45,
            img_w - 0.4,
            img_h - 0.65,
            [
                {
                    "text": f"【{caption}】",
                    "size": 14,
                    "color": WHITE if dark_mode else ATOMY_DARK,
                    "bold": True,
                    "align": "center",
                    "after": 6,
                },
                {
                    "text": hint,
                    "size": 10.5,
                    "color": sub_color,
                    "align": "center",
                    "line": 1.3,
                    "after": 6,
                },
                {
                    "text": "（右键 → 更改图片 → 来自文件，或本地运行脚本自动载入）",
                    "size": 9.5,
                    "color": sub_color,
                    "align": "center",
                },
            ],
            valign="middle",
        )

    # 底部图注栏
    tb(
        slide,
        x + pad + 0.04,
        y + h - cap_h - 0.04,
        w - pad * 2 - 0.08,
        cap_h,
        [
            {
                "runs": [
                    {"text": "●  ", "size": 10, "color": GOLD if dark_mode else ATOMY_BLUE, "bold": True},
                    {"text": caption, "size": 11.5, "color": caption_color, "bold": True},
                    {"text": f"   {hint}", "size": 10, "color": sub_color},
                ]
            }
        ],
        valign="middle",
    )


def resolve_photo_dir(cli_path: str | None) -> Path | None:
    candidates = []
    if cli_path:
        candidates.append(Path(cli_path).expanduser())
    candidates.extend(
        [
            DEFAULT_MAC_PHOTO_DIR,
            Path.cwd() / "5周年旅游照片",
            Path.cwd() / "photos",
            Path(__file__).resolve().parent / "5周年旅游照片",
            Path(__file__).resolve().parent / "photos",
        ]
    )
    for c in candidates:
        if c.exists() and c.is_dir():
            has_imgs = any(p.is_file() and p.suffix.lower() in IMAGE_EXTS for p in c.rglob("*"))
            if has_imgs:
                return c
    return None


def build_ppt(photo_dir: Path | None, output_path: Path) -> int:
    prs = Presentation()
    prs.slide_width = Emu(12192000)   # 13.333 in (16:9)
    prs.slide_height = Emu(6858000)   # 7.5 in
    prs.core_properties.title = "感恩艾多美 · 五载同行携家江南游 —— 5周年员工旅游心得"
    prs.core_properties.subject = "艾多美5周年员工旅游（上海·杭州）全家福与心得体会"

    assigner = PhotoAssigner(photo_dir)
    total_pages = 7

    # =========================================================================
    # 第 1 页：封面 —— 感恩艾多美 · 五载同行 · 携家江南游
    # =========================================================================
    s = new_slide(prs)
    bg(s, ATOMY_DARK)
    rect(s, 0, 0, 0.14, H, ATOMY_BLUE)
    rect(s, 0.14, 0, 0.04, H, GOLD)

    # 顶部福利徽章
    rect(s, 0.82, 0.72, 5.15, 0.38, "163558", "2B5482", rounded=True)
    tb(
        s,
        0.82,
        0.72,
        5.15,
        0.38,
        [
            {
                "text": "ATOMY 艾多美  ·  3 / 5 / 7 / 10 周年员工旅游福利",
                "size": 12,
                "color": "7CCDF4",
                "bold": True,
                "align": "center",
            }
        ],
        valign="middle",
    )

    tb(
        s,
        0.82,
        1.38,
        6.8,
        2.1,
        [
            {"text": "五载同心 · 携爱同行", "size": 42, "color": WHITE, "bold": True, "after": 8},
            {"text": "我的艾多美 5 周年上海 & 杭州家庭之旅", "size": 24, "color": "8ED6FA", "bold": True},
        ],
    )
    rect(s, 0.82, 3.35, 1.35, 0.05, GOLD)

    tb(
        s,
        0.82,
        3.62,
        6.5,
        1.25,
        [
            {
                "text": "在艾多美，每满 3、5、7、10 年都能享受到由公司承担费用的员工旅游。",
                "size": 14.5,
                "color": "E4EEF8",
                "line": 1.38,
                "after": 6,
            },
            {
                "text": "满 5 周年的这一程，我带着父母、老婆和孩子一起走进繁华上海与诗画杭州——把公司的温暖福利，化作一家三代人最珍贵的团圆记忆。",
                "size": 14,
                "color": "B0C7DE",
                "line": 1.38,
            },
        ],
    )

    # 左下行程概览四宫格
    badges = [
        ("里程碑", "入职满 5 周年", "3/5/7/10年专属福利"),
        ("同行家人", "父母 · 老婆 · 孩子", "三代同堂温馨出游"),
        ("双城足迹", "上海  ×  杭州", "外滩豫园 · 西湖雷峰塔"),
        ("舌尖犒赏", "云端自助 & 海鲜大餐", "金茂56层 · 龙虾面包蟹"),
    ]
    bx = 0.82
    for b_tag, b_main, b_sub in badges:
        rect(s, bx, 5.28, 1.56, 1.48, "152F4E", "294B72", rounded=True)
        rect(s, bx, 5.28, 1.56, 0.04, GOLD)
        tb(s, bx + 0.12, 5.42, 1.32, 0.24, [{"text": b_tag, "size": 10.5, "color": "7CCDF4", "bold": True}])
        tb(s, bx + 0.12, 5.72, 1.32, 0.42, [{"text": b_main, "size": 13, "color": WHITE, "bold": True, "line": 1.2}])
        tb(s, bx + 0.12, 6.24, 1.32, 0.36, [{"text": b_sub, "size": 10, "color": "9AB4CE", "line": 1.2}])
        bx += 1.68

    # 右侧封面主照片位（全家福）
    photo_slot(
        s,
        assigner,
        7.65,
        0.72,
        5.0,
        6.04,
        "五周年全家福 · 幸福启程",
        "带着父母、老婆、孩子共赴沪杭之旅",
        keywords=("全家", "合照", "合影", "家人", "父母", "外滩", "西湖", "cover", "01", "1"),
        dark_mode=True,
    )
    notes(
        s,
        "封面页：介绍艾多美每满3、5、7、10年的员工旅游福利（公司承担费用），引出本次满5周年带着父母、老婆、孩子前往上海与杭州的家庭之旅。",
    )

    # =========================================================================
    # 第 2 页：感恩福利与行程总览 —— 公司买单，孝心与陪伴同行
    # =========================================================================
    s = new_slide(prs)
    bg(s, BG_SOFT)
    header(
        s,
        2,
        total_pages,
        "01 · 福利与行程",
        "满 5 年的温暖馈赠：公司承担费用，让孝心与陪伴不留遗憾",
        "每满 3、5、7、10 年员工旅游制度 · 沪杭双城家庭行",
    )

    # 左侧：3/5/7/10年福利阶梯卡片
    rect(s, ML, 1.54, 4.15, 5.38, WHITE, BORDER, rounded=True)
    rect(s, ML, 1.54, 4.15, 0.06, ATOMY_BLUE)
    tb(
        s,
        ML + 0.26,
        1.76,
        3.63,
        0.4,
        [{"text": "艾多美员工旅游福利机制", "size": 17, "color": INK, "bold": True}],
    )
    tb(
        s,
        ML + 0.26,
        2.18,
        3.63,
        0.55,
        [
            {
                "text": "公司为每满 3、5、7、10 年的员工提供专项旅游福利，费用由公司承担，真正把“珍视灵魂”落到实处。",
                "size": 12,
                "color": BODY,
                "line": 1.35,
            }
        ],
    )

    milestones = [
        ("满 3 年", "初绽芳华 · 充电启程", False),
        ("满 5 年", "五载笃行 · 携家沪杭游（本次行程）", True),
        ("满 7 年", "七年深耕 · 远方再叙", False),
        ("满 10 年", "十载同心 · 荣耀盛旅", False),
    ]
    my = 2.88
    for m_yr, m_desc, is_current in milestones:
        fill_c = GOLD_SOFT if is_current else BG_SOFT
        border_c = GOLD if is_current else BORDER
        rect(s, ML + 0.26, my, 3.63, 0.66, fill_c, border_c, line_width=1.2 if is_current else 0.75, rounded=True)
        tb(
            s,
            ML + 0.42,
            my + 0.14,
            1.1,
            0.38,
            [{"text": m_yr, "size": 13.5, "color": GOLD if is_current else ATOMY_BLUE, "bold": True}],
            valign="middle",
        )
        tb(
            s,
            ML + 1.48,
            my + 0.14,
            2.3,
            0.38,
            [{"text": m_desc, "size": 12, "color": INK if is_current else BODY, "bold": is_current}],
            valign="middle",
        )
        my += 0.78

    rect(s, ML + 0.26, 6.08, 3.63, 0.62, ATOMY_LIGHT, rounded=True)
    tb(
        s,
        ML + 0.38,
        6.14,
        3.4,
        0.5,
        [
            {
                "text": "出游成员：我 + 父母 + 老婆 + 孩子（三代同游）\n核心感受：平时忙于工作，这次借公司福利带全家看风景、吃大餐！",
                "size": 11,
                "color": ATOMY_DARK,
                "line": 1.3,
            }
        ],
        valign="middle",
    )

    # 中间：上海 & 杭州双城路线卡片
    stations = [
        (
            "第一站 · 魔都上海（都市繁华与老城烟火）",
            ATOMY_BLUE,
            [
                "• 外滩漫步：饱览万国建筑博览群与黄浦江对岸陆家嘴天际线",
                "• 金茂大厦56层：300元/位高空观景自助餐，陪父母妻儿云端用餐",
                "• 豫园 & 城隍庙：亭台楼阁间感受江南园林与老上海民俗风味",
                "• 南京路步行街：十里洋场霓虹闪烁，全家欢声笑语逛街打卡",
            ],
        ),
        (
            "第二站 · 诗画杭州（湖光山色与海鲜盛宴）",
            GOLD,
            [
                "• 西湖泛舟漫步：水光潋滟晴方好，带父母慢节奏赏湖景",
                "• 雷峰塔登高：俯瞰西湖全景与千年传奇，孩子收获满满",
                "• 丰盛海鲜大餐：大龙虾、膏满肉肥的面包蟹等鲜美海味齐上桌",
            ],
        ),
    ]
    sy = 1.54
    for st_title, st_color, st_items in stations:
        rect(s, 5.05, sy, 4.15, 2.58, WHITE, BORDER, rounded=True)
        rect(s, 5.05, sy, 0.08, 2.58, st_color)
        tb(
            s,
            5.28,
            sy + 0.18,
            3.75,
            0.36,
            [{"text": st_title, "size": 14.5, "color": INK, "bold": True}],
        )
        tb(
            s,
            5.28,
            sy + 0.62,
            3.75,
            1.8,
            [{"text": line, "size": 12, "color": BODY, "line": 1.35, "after": 5} for line in st_items],
        )
        sy += 2.80

    # 右侧：家庭出游照片位
    photo_slot(
        s,
        assigner,
        9.42,
        1.54,
        3.23,
        5.38,
        "三代同堂 · 开心出发",
        "父母身体健康、妻儿笑颜如花",
        keywords=("出发", "高铁", "家庭", "父母", "老婆", "孩子", "全家", "02", "2"),
    )
    notes(s, "第2页：展示艾多美3/5/7/10年员工旅游福利体系，以及本次上海+杭州双城家庭游的整体路线。")

    # =========================================================================
    # 第 3 页：上海站（一）—— 外滩风光 & 金茂大厦56层观景自助
    # =========================================================================
    s = new_slide(prs)
    bg(s, BG_SOFT)
    header(
        s,
        3,
        total_pages,
        "02 · 上海站 · 云端与江景",
        "漫步外滩江风，登顶金茂大厦 56 层享 300 元/位云端观景自助",
        "上海站亮点一：万国建筑群  ×  金茂大厦 56F 高空美食体验",
    )

    # 左侧两个故事卡片
    rect(s, ML, 1.54, 4.55, 2.58, WHITE, BORDER, rounded=True)
    rect(s, ML, 1.54, 4.55, 0.05, ATOMY_BLUE)
    tb(
        s,
        ML + 0.24,
        1.72,
        4.1,
        0.36,
        [{"text": "📍 外滩：浦江两岸的世纪同框", "size": 16, "color": INK, "bold": True}],
    )
    tb(
        s,
        ML + 0.24,
        2.16,
        4.1,
        1.8,
        [
            {
                "text": "带着父母、老婆和孩子走在外滩江堤上，一边是厚重典雅的万国历史建筑群，隔江相望便是高耸入云的陆家嘴摩天大楼。",
                "size": 12.5,
                "color": BODY,
                "line": 1.4,
                "after": 8,
            },
            {
                "text": "父母看着眼前的繁华景象感慨万千，孩子兴奋地指着东方明珠和金茂大厦拍照留念，江风吹拂下满是一家人的温馨。",
                "size": 12.5,
                "color": BODY,
                "line": 1.4,
            },
        ],
    )

    rect(s, ML, 4.32, 4.55, 2.60, WHITE, BORDER, rounded=True)
    rect(s, ML, 4.32, 4.55, 0.05, GOLD)
    tb(
        s,
        ML + 0.24,
        4.50,
        4.1,
        0.36,
        [{"text": "🍽️ 金茂大厦 56 层：300 元/位云端观景自助", "size": 16, "color": INK, "bold": True}],
    )
    tb(
        s,
        ML + 0.24,
        4.94,
        4.1,
        1.8,
        [
            {
                "text": "特意带全家人登上了陆家嘴地标——金茂大厦 56 层，体验 300 元/位的高空观景自助餐。震撼的中庭挑高与窗外城市风光让人眼前一亮。",
                "size": 12.5,
                "color": BODY,
                "line": 1.4,
                "after": 8,
            },
            {
                "text": "因为有公司旅游福利的鼎力支持，可以毫不犹豫地请父母和妻儿享受高品质餐饮。看着长辈吃得舒心、孩子选得开心，内心格外自豪。",
                "size": 12.5,
                "color": BODY,
                "line": 1.4,
            },
        ],
    )

    # 右侧两张大图：外滩 + 金茂56层自助
    photo_slot(
        s,
        assigner,
        5.48,
        1.54,
        3.82,
        5.38,
        "上海外滩 · 浦江留影",
        "与父母、老婆、孩子在外滩江畔合影",
        keywords=("外滩", "东方明珠", "上海", "江", "bund", "03", "3"),
    )
    photo_slot(
        s,
        assigner,
        9.50,
        1.54,
        3.15,
        5.38,
        "金茂大厦 56 层观景自助",
        "300元/位云端美食 · 俯瞰申城风光",
        keywords=("金茂", "56", "自助", "餐厅", "jinmao", "04", "4"),
    )
    notes(s, "第3页：聚焦上海外滩游览与金茂大厦56层300元/位观景自助餐，突出公司承担旅费让员工能带父母妻儿享受高品质体验。")

    # =========================================================================
    # 第 4 页：上海站（二）—— 豫园、城隍庙与南京路步行街
    # =========================================================================
    s = new_slide(prs)
    bg(s, BG_SOFT)
    header(
        s,
        4,
        total_pages,
        "03 · 上海站 · 海派风情",
        "漫游豫园与城隍庙古韵，打卡十里洋场南京路步行街",
        "上海站亮点二：江南古典园林  ×  老城厢烟火气  ×  中华商业第一街",
    )

    # 顶部三栏景点心得卡片
    sh_spots = [
        (
            "01  豫园 · 江南园林雅韵",
            ATOMY_BLUE,
            "九曲桥畔锦鲤戏水，亭台楼阁一步一景。陪着父母慢步在精巧的古典园林里，长辈们对传统建筑赞不绝口。",
        ),
        (
            "02  城隍庙 · 地道老上海烟火",
            GOLD,
            "飞檐翘角下热闹非凡，带家人品尝南翔小笼包与特色传统糕点，热气腾腾的烟火气让一家人胃口大开。",
        ),
        (
            "03  南京路 · 十里洋场繁华夜",
            ATOMY_BLUE,
            "漫步百年南京路步行街，百年老字号与现代潮流交织，叮当车与璀璨霓虹让孩子兴奋不已。",
        ),
    ]
    sx = ML
    for sp_title, sp_color, sp_desc in sh_spots:
        rect(s, sx, 1.50, 3.85, 1.72, WHITE, BORDER, rounded=True)
        rect(s, sx, 1.50, 3.85, 0.05, sp_color)
        tb(s, sx + 0.22, 1.66, 3.4, 0.32, [{"text": sp_title, "size": 15, "color": INK, "bold": True}])
        tb(
            s,
            sx + 0.22,
            2.04,
            3.4,
            1.05,
            [{"text": sp_desc, "size": 12, "color": BODY, "line": 1.35}],
        )
        sx += 4.06

    # 底部三张对应实拍照片位
    photo_slot(
        s,
        assigner,
        ML,
        3.42,
        3.85,
        3.52,
        "豫园 · 亭台九曲桥",
        "古典园林里的全家慢时光",
        keywords=("豫园", "九曲桥", "园林", "yuyuan", "05", "5"),
    )
    photo_slot(
        s,
        assigner,
        ML + 4.06,
        3.42,
        3.85,
        3.52,
        "城隍庙 · 传统民俗与小吃",
        "飞檐古街 · 品尝地道海派风味",
        keywords=("城隍庙", "小吃", "古街", "庙", "06", "6"),
    )
    photo_slot(
        s,
        assigner,
        ML + 8.12,
        3.42,
        3.85,
        3.52,
        "南京路步行街 · 繁华街景",
        "十里洋场打卡留念",
        keywords=("南京路", "步行街", "街", "nanjing", "07", "7"),
    )
    notes(s, "第4页：展示上海豫园、城隍庙、南京路步行街三处经典地标的家庭游览体验与照片。")

    # =========================================================================
    # 第 5 页：杭州站 —— 西湖风光与雷峰塔登高
    # =========================================================================
    s = new_slide(prs)
    bg(s, BG_SOFT)
    header(
        s,
        5,
        total_pages,
        "04 · 杭州站 · 诗画江南",
        "漫步西湖碧波柳浪，登临雷峰塔俯瞰人间天堂",
        "杭州站亮点：西湖十景风韵  ×  雷峰塔登高远眺",
    )

    # 左侧两张大照片（西湖 + 雷峰塔）
    photo_slot(
        s,
        assigner,
        ML,
        1.54,
        3.95,
        5.38,
        "杭州西湖 · 湖光山色",
        "断桥残雪 · 苏堤春晓 · 碧波荡漾",
        keywords=("西湖", "杭州", "湖", "游船", "xihu", "08", "8"),
    )
    photo_slot(
        s,
        assigner,
        ML + 4.12,
        1.54,
        3.95,
        5.38,
        "雷峰塔 · 登高望远",
        "登塔俯瞰西湖全景与古今传奇",
        keywords=("雷峰塔", "塔", "leifeng", "09", "9"),
    )

    # 右侧：杭州游记与家庭感悟
    rect(s, 8.95, 1.54, 3.70, 5.38, WHITE, BORDER, rounded=True)
    rect(s, 8.95, 1.54, 3.70, 0.06, ATOMY_BLUE)

    tb(
        s,
        9.20,
        1.78,
        3.20,
        0.38,
        [{"text": "🌿 江南杭州 · 身心放松之旅", "size": 16.5, "color": INK, "bold": True}],
    )
    hz_points = [
        (
            "西湖畔的慢步调",
            "告别了魔都上海的快节奏，来到杭州西湖边，迎面是微风拂柳与碧波游船。陪着父母沿着湖畔漫步，节奏舒缓惬意，最适合长辈放松身心。",
        ),
        (
            "登雷峰塔揽全湖之胜",
            "全家一起登上雷峰塔（景区贴心的扶梯与电梯让父母毫无爬山压力）。站在塔顶极目远眺，整座西湖的山水画卷尽收眼底，孩子还兴致勃勃地听起了白娘子的传说。",
        ),
        (
            "三代同游的温情时刻",
            "看着父母在西湖边舒展的笑容、老婆和孩子并肩看风景的背影，深刻感受到：最美的风景不仅是西湖山水，更是身旁整整齐齐的家人。",
        ),
    ]
    hy = 2.30
    for h_title, h_desc in hz_points:
        rect(s, 9.20, hy, 3.20, 1.40, BG_SOFT, BORDER, rounded=True)
        tb(
            s,
            9.36,
            hy + 0.12,
            2.90,
            0.28,
            [{"text": f"◆ {h_title}", "size": 13, "color": ATOMY_BLUE, "bold": True}],
        )
        tb(
            s,
            9.36,
            hy + 0.42,
            2.90,
            0.90,
            [{"text": h_desc, "size": 11, "color": BODY, "line": 1.32}],
        )
        hy += 1.52

    notes(s, "第5页：展示杭州西湖与雷峰塔的游览体验，体现长辈轻松舒心、孩子增长见识的家庭温馨。")

    # =========================================================================
    # 第 6 页：舌尖上的幸福 —— 龙虾、面包蟹等海鲜大餐
    # =========================================================================
    s = new_slide(prs)
    bg(s, BG_SOFT)
    header(
        s,
        6,
        total_pages,
        "05 · 舌尖盛宴 · 海鲜大餐",
        "大龙虾、膏满面包蟹齐上桌：用丰盛海鲜大餐犒赏最爱的家人",
        "美食记录：金茂 56 层云端自助之外，更有一家人围坐共享的澎湃海鲜盛宴",
    )

    # 左侧：海鲜大餐特色描述卡
    rect(s, ML, 1.54, 4.45, 5.38, WHITE, BORDER, rounded=True)
    rect(s, ML, 1.54, 4.45, 0.06, GOLD)

    tb(
        s,
        ML + 0.26,
        1.78,
        3.95,
        0.4,
        [{"text": "🦞 鲜味满桌 · 团圆欢聚", "size": 17.5, "color": INK, "bold": True}],
    )
    tb(
        s,
        ML + 0.26,
        2.26,
        3.95,
        0.85,
        [
            {
                "text": "除了上海金茂大厦 56 层 300 元/位的观景自助，这次旅途中最让全家人念念不忘的，还有丰盛过瘾的海鲜大餐！",
                "size": 12.5,
                "color": BODY,
                "line": 1.38,
            }
        ],
    )

    dishes = [
        ("招牌大龙虾", "肉质紧实弹牙、鲜甜饱满，端上桌瞬间成为全场焦点，孩子和父母都赞不绝口。"),
        ("爆膏面包蟹", "蟹壳里满满的红膏与蟹黄，蟹肉厚实鲜美，一口下去全是满足感。"),
        ("多款时令海味", "各类鲜虾贝类琳琅满目，一家人围坐举杯，欢声笑语间洗去旅途疲惫。"),
    ]
    dy = 3.22
    for d_name, d_desc in dishes:
        rect(s, ML + 0.26, dy, 3.93, 0.96, GOLD_SOFT, "F0DEC2", rounded=True)
        tb(
            s,
            ML + 0.42,
            dy + 0.10,
            3.6,
            0.28,
            [{"text": f"★  {d_name}", "size": 13, "color": GOLD, "bold": True}],
        )
        tb(
            s,
            ML + 0.42,
            dy + 0.40,
            3.6,
            0.48,
            [{"text": d_desc, "size": 11.2, "color": BODY, "line": 1.3}],
        )
        dy += 1.08

    rect(s, ML + 0.26, 6.46, 3.93, 0.34, ATOMY_LIGHT, rounded=True)
    tb(
        s,
        ML + 0.26,
        6.46,
        3.93,
        0.34,
        [{"text": "感谢公司报销旅游费用，让家人吃得尽兴、玩得舒心！", "size": 11, "color": ATOMY_BLUE, "bold": True, "align": "center"}],
        valign="middle",
    )

    # 右侧：两张海鲜美食照片位
    photo_slot(
        s,
        assigner,
        5.38,
        1.54,
        3.62,
        5.38,
        "招牌大龙虾 · 鲜美上桌",
        "个头饱满 · 肉质Q弹鲜甜",
        keywords=("龙虾", "虾", "海鲜", "lobster", "seafood", "10"),
    )
    photo_slot(
        s,
        assigner,
        9.20,
        1.54,
        3.45,
        5.38,
        "爆膏面包蟹 & 海鲜盛宴",
        "膏满黄肥 · 全家围坐大快朵颐",
        keywords=("蟹", "面包蟹", "大餐", "美食", "吃", "crab", "11"),
    )
    notes(s, "第6页：重点展示旅途中的海鲜大餐（大龙虾、面包蟹等），体现一家人共享美食的幸福时刻。")

    # =========================================================================
    # 第 7 页：旅游心得总结 —— 感恩艾多美，奔赴下一个7年、10年
    # =========================================================================
    s = new_slide(prs)
    bg(s, BG_SOFT)
    header(
        s,
        7,
        total_pages,
        "06 · 旅游心得与感恩",
        "一份福利温暖一家人：感恩艾多美五载培养，满电再出发！",
        "五周年员工旅游总结  ·  向着 7 周年、10 周年继续笃行",
    )

    reflections = [
        (
            "01  感恩公司：有温度的福利制度",
            ATOMY_BLUE,
            "艾多美每满 3、5、7、10 年的公费旅游福利，不仅是对员工多年付出的认可，更是对员工背后家庭的关怀。因为有公司的承担，这次五周年之旅才能毫无后顾之忧地给家人安排最好的体验。",
        ),
        (
            "02  珍视家人：最好的爱是高质量陪伴",
            GOLD,
            "平时忙于工作，很少有机会同时带父母、老婆和孩子出远门。从外滩江风、金茂 56 层自助、豫园城隍庙，到西湖雷峰塔与海鲜大餐，看到父母的笑容与孩子的成长，所有的忙碌都有了意义。",
        ),
        (
            "03  笃定前行：满电回归，奔赴 7 年与 10 年",
            ATOMY_DARK,
            "身心在江南山水中得到了彻底的放松与充电，家人的支持与认可更化作我前进的动力。回到岗位后，我将以更饱满的热情投入工作，与艾多美一起成长，期待 7 周年、10 周年的下一次约定！",
        ),
    ]

    ry = 1.54
    for r_title, r_color, r_body in reflections:
        rect(s, ML, ry, 7.25, 1.62, WHITE, BORDER, rounded=True)
        rect(s, ML, ry, 0.09, 1.62, r_color)
        tb(
            s,
            ML + 0.28,
            ry + 0.16,
            6.75,
            0.34,
            [{"text": r_title, "size": 15.5, "color": INK, "bold": True}],
        )
        tb(
            s,
            ML + 0.28,
            ry + 0.56,
            6.75,
            0.92,
            [{"text": r_body, "size": 12, "color": BODY, "line": 1.38}],
        )
        ry += 1.84

    # 右侧：温馨收尾照片位 + 结语横幅
    photo_slot(
        s,
        assigner,
        8.18,
        1.54,
        4.47,
        4.15,
        "旅途美好定格 · 幸福全家福",
        "感恩艾多美 · 让工作与家庭生活同样精彩",
        keywords=("合影", "全家", "笑", "家庭", "最后", "end", "12"),
    )

    rect(s, 8.18, 5.86, 4.47, 1.02, ATOMY_DARK, rounded=True)
    rect(s, 8.18, 5.86, 4.47, 0.05, GOLD)
    tb(
        s,
        8.36,
        5.96,
        4.11,
        0.84,
        [
            {
                "text": "“因为艾多美，我能带着最爱的人看更大的世界。”",
                "size": 13,
                "color": WHITE,
                "bold": True,
                "align": "center",
                "after": 4,
            },
            {
                "text": "感谢公司 · 感谢领导与伙伴 · 下一站 7 周年见！",
                "size": 11.5,
                "color": "8ED6FA",
                "align": "center",
            },
        ],
        valign="middle",
    )
    notes(s, "第7页：总结五周年员工旅游的三点核心心得（感恩公司福利、珍惜家人陪伴、满电回归奔赴7年与10年）。")

    prs.save(str(output_path))
    return len(assigner.all_photos)


def main() -> None:
    parser = argparse.ArgumentParser(description="生成艾多美5周年员工旅游心得PPT（自动配图）")
    parser.add_argument(
        "-p",
        "--photos",
        default=None,
        help="照片文件夹路径（默认自动查找 /Users/jiangxiaoyi/sam's document/公司旅游心得/5周年旅游照片）",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="艾多美5周年员工旅游心得_上海杭州.pptx",
        help="输出的 PPTX 文件路径",
    )
    args = parser.parse_args()

    photo_dir = resolve_photo_dir(args.photos)
    out_path = Path(args.output)
    count = build_ppt(photo_dir, out_path)
    if photo_dir:
        print(f"已从照片目录 [{photo_dir}] 扫描到 {count} 张照片并自动排版生成：{out_path}")
    else:
        print(
            f"已生成 PPT：{out_path}（当前环境未检测到本地照片目录，已预留高清实拍相框位；"
            f"在您的 Mac 上运行 `python3 build_trip_ppt.py` 即可自动从 `{DEFAULT_MAC_PHOTO_DIR}` 导入全部照片）"
        )


if __name__ == "__main__":
    main()
