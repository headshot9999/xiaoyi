#!/usr/bin/env python3
"""艾多美一百兆论证 · 演示文稿"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SANS = "Noto Sans CJK SC"
SANS_M = "Noto Sans CJK SC Medium"

WHITE = "FFFFFF"
INK = "121212"
SUB = "5E5E5E"
SOFT = "F4F4F4"
LINE = "E6E6E6"
RED = "D0121A"
DARK = "111111"
GRAY = "8A8A8A"
ON_DARK = "F3F3F3"
ON_DARK_DIM = "A0A0A0"

W = 13.333333
H = 7.5
ML = 0.62


def rgb(h):
    return RGBColor.from_string(h)


def _anchor(tf, valign):
    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    bodyPr.set("anchor", {"top": "t", "middle": "ctr", "bottom": "b"}[valign])
    for key in ("lIns", "rIns", "tIns", "bIns"):
        bodyPr.set(key, "0")


def _run(p, text, font, size, color):
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = False
    run.font.italic = False
    run.font.color.rgb = rgb(color)
    run.font.name = font
    rPr = run._r.get_or_add_rPr()
    rPr.set("lang", "zh-CN")
    rPr.set("altLang", "en-US")
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = rPr.makeelement(qn(f"a:{tag}"), {"typeface": font})
            rPr.append(el)
        else:
            el.set("typeface", font)
    return run


def tb(slide, x, y, w, h, paragraphs, valign="top"):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    _anchor(tf, valign)
    align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    for i, info in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align_map[info.get("align", "left")]
        if "line" in info:
            p.line_spacing = info["line"]
        if info.get("before"):
            p.space_before = Pt(info["before"])
        if info.get("after"):
            p.space_after = Pt(info["after"])
        runs = info.get("runs")
        if runs:
            for r in runs:
                _run(p, r["text"], r.get("font", SANS), r["size"], r["color"])
        else:
            _run(p, info["text"], info.get("font", SANS), info["size"], info["color"])
    return shape


def rect(slide, x, y, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = rgb(fill)
    sh.line.fill.background()
    return sh


def oval(slide, x, y, s, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(s), Inches(s))
    sh.fill.solid()
    sh.fill.fore_color.rgb = rgb(fill)
    sh.line.fill.background()
    return sh


def bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def new(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def page(slide, n, total, dark=False):
    tb(
        slide,
        11.15,
        7.08,
        1.55,
        0.26,
        [{"text": f"{n:02d}  /  {total:02d}", "font": SANS, "size": 11, "color": "6E6E6E" if dark else "B0B0B0", "align": "right"}],
    )


def header(slide, section, title):
    rect(slide, 0, 0, W, 0.08, RED)
    tb(slide, ML, 0.32, 8, 0.28, [{"text": section, "font": SANS, "size": 13, "color": RED}])
    tb(slide, ML, 0.64, 12, 0.58, [{"text": title, "font": SANS_M, "size": 32, "color": INK}])


def build():
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    prs.core_properties.title = "艾多美何以成为100兆韩元企业"
    prs.core_properties.subject = "蒙想讯息 · 2026年9月22日"
    total = 15

    # 01 cover
    s = new(prs)
    bg(s, DARK)
    rect(s, 0, 0, 0.14, H, RED)
    tb(s, 0.85, 0.72, 10, 0.32, [{"text": "蒙想讯息  ·  2026年9月22日", "font": SANS, "size": 15, "color": ON_DARK_DIM}])
    tb(
        s,
        0.82,
        2.15,
        11.5,
        2.15,
        [
            {"text": "艾多美何以成为", "font": SANS_M, "size": 48, "color": WHITE, "after": 6},
            {"text": "一百兆韩元企业", "font": SANS_M, "size": 48, "color": WHITE},
        ],
    )
    rect(s, 0.85, 4.6, 1.35, 0.045, RED)
    tb(s, 0.85, 4.85, 8, 0.4, [{"text": "把口号写成结构", "font": SANS, "size": 20, "color": "D0D0D0"}])
    page(s, 1, total, dark=True)
    notes(s, "停一拍。下一页直接给两个数字。")

    # 02 tenfold
    s = new(prs)
    bg(s, WHITE)
    rect(s, 0, 0, 6.35, H, DARK)
    rect(s, 6.35, 0, 0.08, H, RED)
    tb(s, 0.7, 1.85, 5, 0.32, [{"text": "眼下的第一", "font": SANS, "size": 15, "color": ON_DARK_DIM}])
    tb(s, 0.66, 2.3, 5.3, 1.25, [{"text": "十兆", "font": SANS_M, "size": 76, "color": WHITE}])
    tb(s, 0.7, 3.75, 5, 0.4, [{"text": "大约六十年", "font": SANS_M, "size": 22, "color": ON_DARK}])
    tb(s, 0.7, 4.35, 5, 0.8, [{"text": "安利。行业里走在前面的那个数字。", "font": SANS, "size": 16, "color": ON_DARK_DIM, "line": 1.4}])
    tb(s, 7.05, 1.85, 5.5, 0.32, [{"text": "要去的地方", "font": SANS, "size": 15, "color": RED}])
    tb(s, 7.0, 2.3, 5.8, 1.25, [{"text": "一百兆", "font": SANS_M, "size": 72, "color": RED}])
    tb(s, 7.05, 3.75, 5.5, 0.4, [{"text": "大约十五年", "font": SANS_M, "size": 22, "color": INK}])
    tb(s, 7.05, 4.35, 5.5, 0.85, [{"text": "蒙想还在的年月。不是一百年之后。", "font": SANS, "size": 16, "color": SUB, "line": 1.4}])
    tb(s, 7.05, 6.15, 5.6, 0.7, [{"text": "再往后，是千兆。\n要比过的，是沃尔玛，是亚马逊。", "font": SANS, "size": 16, "color": INK, "line": 1.45}])
    page(s, 2, total)
    notes(s, "十兆用了大约六十年。一百兆要在大约十五年里走到。后面每一页，都是写给不信这件事的人。")

    # 03 standard
    s = new(prs)
    bg(s, WHITE)
    header(s, "标准", "信念握得住口号")
    rect(s, ML, 1.85, 5.85, 4.55, SOFT)
    rect(s, 6.85, 1.85, 5.85, 4.55, DARK)
    tb(s, 0.95, 2.15, 5.2, 0.55, [{"text": "信念", "font": SANS_M, "size": 28, "color": INK}])
    tb(
        s,
        0.95,
        3.0,
        5.15,
        2.6,
        [{"text": "中心是自己。\n喊一喊，念头会更紧。\n一百兆不会因此发生。", "font": SANS, "size": 20, "color": SUB, "line": 1.55}],
    )
    tb(s, 7.18, 2.15, 5.2, 0.55, [{"text": "结构", "font": SANS_M, "size": 28, "color": WHITE}])
    tb(
        s,
        7.18,
        3.0,
        5.15,
        2.6,
        [{"text": "中心是事情本身。\n信，或者不信，它都往那里去。\n论证要写给不信的人。", "font": SANS, "size": 20, "color": "E4E4E4", "line": 1.55}],
    )
    page(s, 3, total)
    notes(s, "信念是把自己的想法握紧。要交的是结构。信与不信，都会看见它往那里去。")

    # 04 who
    s = new(prs)
    bg(s, WHITE)
    header(s, "战场", "先把人分开")
    tb(
        s,
        ML,
        1.5,
        12,
        0.4,
        [{"text": "场子铺开，是让人自己挑。教科书接着说，高低都做。方向会因此松掉。", "font": SANS, "size": 16, "color": SUB}],
    )
    rows = [
        ("01", "丢掉", "最上约 1%", "嫌便宜。东西再好，也不用。", False),
        ("02", "丢掉", "只认便宜", "韩国大约三成。购买力更弱的地方，可以到七成。", False),
        ("03", "留下", "绝对品质，绝对价格", "这一层拿下一半，千兆也撑得住。", True),
    ]
    y = 2.15
    for num, verb, head, body, on in rows:
        rect(s, ML, y, 12.08, 1.38, SOFT if on else WHITE)
        if not on:
            rect(s, ML, y + 1.36, 12.08, 0.015, LINE)
        else:
            rect(s, ML, y, 0.08, 1.38, RED)
        tb(s, 0.95, y + 0.42, 0.7, 0.45, [{"text": num, "font": SANS_M, "size": 18, "color": RED}])
        tb(s, 1.75, y + 0.44, 1.2, 0.4, [{"text": verb, "font": SANS, "size": 16, "color": SUB}])
        tb(s, 3.15, y + 0.38, 4.6, 0.55, [{"text": head, "font": SANS_M, "size": 22, "color": INK}])
        tb(s, 7.9, y + 0.42, 4.5, 0.55, [{"text": body, "font": SANS, "size": 15, "color": SUB}])
        y += 1.5
    page(s, 4, total)
    notes(s, "不要全部市场。丢掉最上约百分之一，也丢掉只求最贱的那一层。一百兆的分母，是同时要绝对品质和绝对价格的人，再取一半。")

    # 05 35
    s = new(prs)
    bg(s, WHITE)
    header(s, "津贴", "三十五个点，不能向顾客要")
    rect(s, ML, 1.7, 4.35, 4.7, SOFT)
    tb(s, 0.9, 2.15, 4, 1.3, [{"text": "35%", "font": SANS_M, "size": 80, "color": RED}])
    tb(
        s,
        0.95,
        3.7,
        3.7,
        2.1,
        [{"text": "会员侧的津贴。\n早年靠少了中间商。\n现在的渠道，大多只过一手。", "font": SANS, "size": 16, "color": SUB, "line": 1.5}],
    )
    comps = [
        ("好市多", "毛利为零。年费约八万韩元。买一千万与买一亿，都是这一笔。生产者不另加负担。"),
        ("易买得", "进货比重常在六成到七成。"),
        ("电视购物", "播出费用大约四成二，生产者留下五成八。"),
    ]
    y = 1.75
    for i, (name, desc) in enumerate(comps):
        tb(s, 5.4, y, 7.2, 0.4, [{"text": name, "font": SANS_M, "size": 20, "color": INK}])
        tb(s, 5.4, y + 0.48, 7.2, 0.85, [{"text": desc, "font": SANS, "size": 15, "color": SUB, "line": 1.35}])
        y += 1.55
        if i < 2:
            rect(s, 5.4, y - 0.18, 7.15, 0.015, LINE)
    page(s, 5, total)
    notes(s, "向消费者加价，等于在好市多面前认输。对方降不下来的成本，渠道也降不下来。")

    # 06 source
    s = new(prs)
    bg(s, WHITE)
    header(s, "水源", "三件事要同时成立")
    cards = [
        ("01", "会员", "大约半价"),
        ("02", "事业者", "津贴还在"),
        ("03", "公司", "仍有余量"),
    ]
    x = ML
    for num, who, what in cards:
        rect(s, x, 1.75, 3.9, 3.15, SOFT)
        rect(s, x, 1.75, 3.9, 0.08, RED)
        tb(s, x + 0.32, 2.1, 3.2, 0.35, [{"text": num, "font": SANS_M, "size": 14, "color": RED}])
        tb(s, x + 0.32, 2.6, 3.2, 0.4, [{"text": who, "font": SANS, "size": 16, "color": SUB}])
        tb(s, x + 0.32, 3.15, 3.2, 0.9, [{"text": what, "font": SANS_M, "size": 32, "color": INK}])
        x += 4.15
    tb(
        s,
        ML,
        5.25,
        12,
        1.2,
        [{"text": "这三十五个点的水源，在生产工序里。\n不是加在价格上，也不是从合作方的利润里硬削下来。", "font": SANS, "size": 18, "color": INK, "line": 1.5}],
    )
    page(s, 6, total)
    notes(s, "半价、津贴、公司余量，靠的是同一处水源。下一页用牙刷把算法拆开。")

    # 07 toothbrush
    s = new(prs)
    bg(s, WHITE)
    header(s, "牙刷", "差出来的，不是毛利")
    tb(s, ML, 1.6, 6, 1.15, [{"text": "1,200", "font": SANS_M, "size": 72, "color": RED}])
    tb(s, ML, 2.9, 5.5, 0.4, [{"text": "韩元", "font": SANS, "size": 16, "color": SUB}])
    tb(s, ML, 3.4, 5.5, 0.7, [{"text": "二十年，没有人跟上。", "font": SANS_M, "size": 18, "color": INK}])
    rect(s, 6.7, 1.65, 6.0, 3.55, SOFT)
    tb(s, 7.0, 1.85, 5.4, 0.3, [{"text": "市价", "font": SANS, "size": 14, "color": SUB}])
    tb(s, 7.0, 2.2, 5.4, 0.7, [{"text": "2,000 – 3,000", "font": SANS_M, "size": 32, "color": INK}])
    rect(s, 7.0, 3.1, 5.4, 0.1, "E4E4E4")
    rect(s, 7.0, 3.1, 2.16, 0.1, RED)
    tb(s, 7.0, 3.35, 5.4, 0.3, [{"text": "差价里叠着三层", "font": SANS, "size": 13, "color": SUB}])
    layers = [("制造", "型号一多，每样只做一点"), ("信任", "广告，和店铺"), ("选择", "挑选，买错，退货")]
    yy = 3.8
    for name, desc in layers:
        tb(s, 7.0, yy, 1.15, 0.35, [{"text": name, "font": SANS_M, "size": 15, "color": RED}])
        tb(s, 8.2, yy, 4.1, 0.35, [{"text": desc, "font": SANS, "size": 15, "color": INK}])
        yy += 0.4
    tb(s, ML, 5.6, 12, 0.8, [{"text": "一个造型。腔数做大，原料买吨袋，货款付现金。\n策展若停在挑一件现成的好货，别人明天就能做。", "font": SANS, "size": 16, "color": SUB, "line": 1.45}])
    page(s, 7, total)
    notes(s, "人的嘴不变，模具就不改。把产量聚到一个造型上，成本才下得去。")

    # 08 HemoHIM
    s = new(prs)
    bg(s, WHITE)
    header(s, "HemoHIM", "草根煮一煮，为什么那么贵")
    rect(s, ML, 1.6, 5.7, 2.15, SOFT)
    rect(s, 6.95, 1.6, 5.75, 2.15, "1A1A1A")
    tb(s, 0.9, 1.78, 5.2, 0.3, [{"text": "原先", "font": SANS, "size": 14, "color": SUB}])
    tb(s, 0.88, 2.15, 5.2, 0.9, [{"text": "770,000", "font": SANS_M, "size": 40, "color": INK}])
    tb(s, 0.9, 3.15, 5.2, 0.35, [{"text": "韩元", "font": SANS, "size": 14, "color": SUB}])
    tb(s, 7.22, 1.78, 5.2, 0.3, [{"text": "现在", "font": SANS, "size": 14, "color": "E08A8A"}])
    tb(s, 7.2, 2.1, 5.2, 1.0, [{"text": "76,500", "font": SANS_M, "size": 44, "color": WHITE}])
    tb(s, 7.22, 3.15, 5.2, 0.35, [{"text": "韩元", "font": SANS, "size": 14, "color": "E08A8A"}])
    tb(
        s,
        ML,
        3.95,
        12,
        0.4,
        [{"text": "价格停在十年前。品质走到十年后。依据是一个月十万盒。", "font": SANS, "size": 16, "color": INK}],
    )
    facts = [
        ("0.5  →  18 吨", "三十六倍。操作的人仍是一个。"),
        ("干燥与切断", "省掉。鲜品洗净，直接取汁。"),
        ("合同栽培", "药材没有弹性。锁在涨价之前。"),
    ]
    x = ML
    for head, body in facts:
        rect(s, x, 4.6, 3.9, 1.85, SOFT)
        tb(s, x + 0.28, 4.8, 3.4, 0.7, [{"text": head, "font": SANS_M, "size": 18, "color": INK}])
        tb(s, x + 0.28, 5.5, 3.4, 0.7, [{"text": body, "font": SANS, "size": 14, "color": SUB, "line": 1.35}])
        x += 4.15
    page(s, 8, total)
    notes(s, "罐子从零点五吨做到十八吨。药房要烘干、切断；大罐用鲜品，这两道就没有了。")

    # 09 process
    s = new(prs)
    bg(s, WHITE)
    header(s, "一贯工程", "重做的是整条线")
    cols = [
        ("01", "走到原料", "调配不是源头。温度、压力、配比，向原料厂要。产线怎么转，图纸在设备厂手里。"),
        ("02", "先锁住，再放量", "工业品可以连夜加产。药材要一两年。市场做大之前，工序先放到不能再省。"),
        ("03", "会做，才交出运转", "设备我们出，原料我们买，对方只运转。一起，可以把小公司做到一百亿。不一起，自己做。"),
    ]
    x = ML
    for num, head, body in cols:
        tb(s, x, 1.7, 3.7, 0.4, [{"text": num, "font": SANS_M, "size": 16, "color": RED}])
        tb(s, x, 2.2, 3.75, 1.0, [{"text": head, "font": SANS_M, "size": 24, "color": INK}])
        tb(s, x, 3.4, 3.7, 2.2, [{"text": body, "font": SANS, "size": 16, "color": SUB, "line": 1.5}])
        x += 4.15
    rect(s, ML, 5.85, 12.08, 0.015, LINE)
    tb(s, ML, 6.1, 12, 0.5, [{"text": "绝对价格，是谁来做都无法更便宜。跟，就一起。不跟，自己做。", "font": SANS_M, "size": 16, "color": INK}])
    page(s, 9, total)
    notes(s, "一百兆不是等对方肯配合。优化不停。")

    # 10 daiso
    s = new(prs)
    bg(s, WHITE)
    header(s, "对手", "生产侧，对方已经走到了")
    rect(s, ML, 1.7, 5.9, 3.35, SOFT)
    rect(s, 6.8, 1.7, 5.9, 3.35, SOFT)
    tb(s, 0.95, 1.95, 5.2, 0.32, [{"text": "大创的销售", "font": SANS, "size": 15, "color": SUB}])
    tb(s, 0.92, 2.4, 5.3, 1.05, [{"text": "4.5 兆", "font": SANS_M, "size": 56, "color": INK}])
    tb(s, 0.95, 3.65, 5.2, 0.9, [{"text": "增长维持在三成以上。\n策展、放量、进产线，它都做。", "font": SANS, "size": 16, "color": SUB, "line": 1.4}])
    tb(s, 7.12, 1.95, 5.2, 0.32, [{"text": "艾多美眼下的增长", "font": SANS, "size": 15, "color": RED}])
    tb(s, 7.08, 2.4, 5.3, 1.05, [{"text": "约 3%", "font": SANS_M, "size": 56, "color": RED}])
    tb(s, 7.12, 3.65, 5.2, 0.9, [{"text": "品质更好、设计更好，\n解释不了这一仗。", "font": SANS, "size": 16, "color": SUB, "line": 1.4}])
    tb(
        s,
        ML,
        5.35,
        12.1,
        1.2,
        [{"text": "便宜不等于粗糙。生产上能挖的，对方大多已经挖到了。\n差别要到销售这一侧去找。", "font": SANS, "size": 18, "color": INK, "line": 1.45}],
    )
    page(s, 10, total)
    notes(s, "大创比一般卖场更要认真看。说品牌更高，是不够的。")

    # 11 organization
    s = new(prs)
    bg(s, WHITE)
    header(s, "组织", "店不存在，队伍先走")
    rect(s, ML, 1.6, 5.9, 2.55, SOFT)
    rect(s, 6.8, 1.6, 5.9, 2.55, "1A1A1A")
    tb(s, 0.92, 1.8, 5.2, 0.4, [{"text": "大创", "font": SANS_M, "size": 22, "color": INK}])
    tb(s, 0.92, 2.35, 5.25, 1.5, [{"text": "必须先有店，人再来。\n店的成本，和到店的成本，都是实的。", "font": SANS, "size": 16, "color": SUB, "line": 1.45}])
    tb(s, 7.1, 1.8, 5.2, 0.4, [{"text": "艾多美", "font": SANS_M, "size": 22, "color": WHITE}])
    tb(s, 7.1, 2.35, 5.25, 1.5, [{"text": "没有店。人自己用，也介绍。\n半价的体验代替广告。省下的，就是那 35%。", "font": SANS, "size": 16, "color": "E6E6E6", "line": 1.45}])
    steps = ["绝对价格", "会员", "产量", "成本再降"]
    xs = [1.15, 4.25, 7.35, 10.45]
    rect(s, xs[0] + 0.22, 4.72, xs[-1] - xs[0], 0.035, RED)
    for i, (name, x) in enumerate(zip(steps, xs)):
        oval(s, x, 4.52, 0.42, RED if i == 0 else DARK)
        tb(s, x - 0.55, 5.1, 1.55, 0.4, [{"text": name, "font": SANS_M, "size": 14, "color": INK, "align": "center"}])
    tb(
        s,
        ML,
        5.7,
        12.1,
        1.05,
        [{"text": "摩洛哥，会员十万以上。孟加拉，约二十万。收入差这么远，卖的是同一套产品。\n轮子没有刹车。邻里若觉得街上有差不多的东西、价格低很多，话就传不出去。", "font": SANS, "size": 15, "color": SUB, "line": 1.45}],
    )
    page(s, 11, total)
    notes(s, "组织先于店铺移动。绝对价格招来会员，会员堆出产量，产量把成本再压低。")

    # 12 curation
    s = new(prs)
    bg(s, WHITE)
    header(s, "策展", "不停在货架上")
    qs = [
        ("01", "留下什么", "一个品类，一个。品质不拿来换价格。顾客不必在几百个相近的东西之间比较。"),
        ("02", "如何守住", "从原物到流通。工厂不必座座自有。原料、设备基准和数据，要在自己手里。"),
        ("03", "对准谁", "同一种产品，不再对所有人说同一句话。身体不同，建议就不同。"),
    ]
    y = 1.6
    for num, head, body in qs:
        tb(s, ML, y, 0.7, 0.45, [{"text": num, "font": SANS_M, "size": 18, "color": RED}])
        tb(s, 1.5, y, 3.3, 0.5, [{"text": head, "font": SANS_M, "size": 24, "color": INK}])
        tb(s, 5.1, y + 0.05, 7.5, 0.7, [{"text": body, "font": SANS, "size": 16, "color": SUB, "line": 1.35}])
        y += 1.25
        if y < 5.2:
            rect(s, ML, y - 0.22, 12.08, 0.015, LINE)
    rect(s, ML, 5.45, 12.08, 1.15, SOFT)
    tb(
        s,
        0.9,
        5.7,
        11.5,
        0.7,
        [{"text": "五百个，不能按五百个算。一把牙刷，顶别人那几百款里被挑剩的一个。", "font": SANS_M, "size": 18, "color": INK}],
    )
    page(s, 12, total)
    notes(s, "大量生产人人懂。进不去，是因为两边都在躲风险。愿景要先被看见，循环才开始。")

    # 13 3A
    s = new(prs)
    bg(s, WHITE)
    header(s, "三 A", "三台引擎，一个结构")
    engines = [
        ("APP", "个人平台", "自己的店，自己的健康，自己的教育，收成一处。一小时，变成五分钟。"),
        ("A-Care", "艾护理", "从“请用”，进到血、饮食，和以后的十年。越用越准，人就留下来。"),
        ("AZA", "阿扎", "鸡蛋、油、加油，接进同一处。十万个品目负责人每天来。一百兆，靠大约五百个主力品目。"),
    ]
    x = ML
    for en, cn, body in engines:
        rect(s, x, 1.65, 3.9, 3.7, SOFT)
        rect(s, x, 1.65, 3.9, 0.08, RED)
        tb(s, x + 0.3, 1.95, 3.3, 0.35, [{"text": en, "font": SANS_M, "size": 14, "color": RED}])
        tb(s, x + 0.3, 2.4, 3.3, 0.55, [{"text": cn, "font": SANS_M, "size": 26, "color": INK}])
        tb(s, x + 0.3, 3.2, 3.3, 1.8, [{"text": body, "font": SANS, "size": 15, "color": SUB, "line": 1.45}])
        x += 4.15
    tb(
        s,
        ML,
        5.6,
        12.1,
        0.9,
        [{"text": "事业者从推销变成顾问，变成这家平台的主人。\n消费、推荐和组织，变成留得下的资产。", "font": SANS, "size": 16, "color": INK, "line": 1.45}],
    )
    page(s, 13, total)
    notes(s, "三 A 不是三款应用。艾护理负责关系，阿扎负责每天来，个人平台负责把事业做成方法。")

    # 14 four
    s = new(prs)
    bg(s, WHITE)
    header(s, "结论", "四件事同时为真")
    cells = [
        ("01", "这条链不容易被照抄", "抄一张价格表容易。抄育苗、栽培、图纸和数据，要有愿景，也要有敢先投入的人。"),
        ("02", "不附带对方肯配合", "跟，就一起把小公司做大。不跟，优化也不停。"),
        ("03", "一半，就是千兆的底", "不要全部市场。绝对品质、绝对价格这一层的一半，已经够。"),
        ("04", "人留下，产量才留下", "健康越用越准，日常开支接在一处，内容和组织长在自己的平台上。"),
    ]
    positions = [(ML, 1.6), (6.85, 1.6), (ML, 4.15), (6.85, 4.15)]
    for (num, head, body), (x, y) in zip(cells, positions):
        rect(s, x, y, 5.85, 2.3, SOFT)
        tb(s, x + 0.32, y + 0.25, 1.0, 0.4, [{"text": num, "font": SANS_M, "size": 16, "color": RED}])
        tb(s, x + 0.32, y + 0.7, 5.2, 0.5, [{"text": head, "font": SANS_M, "size": 20, "color": INK}])
        tb(s, x + 0.32, y + 1.28, 5.2, 0.8, [{"text": body, "font": SANS, "size": 14, "color": SUB, "line": 1.35}])
    page(s, 14, total)
    notes(s, "一百兆是这四句同时转起来的结果。可以分开讲，不能分开成立。")

    # 15 close
    s = new(prs)
    bg(s, DARK)
    rect(s, 0, 0, 0.14, H, RED)
    lines = ["生产被重做", "消费者被组织", "选择被拿掉", "三十五个点有了出处"]
    y = 0.85
    for line in lines:
        tb(s, 0.85, y, 11, 0.58, [{"text": line, "font": SANS_M, "size": 28, "color": WHITE}])
        y += 0.72
    tb(s, 0.85, 4.15, 11, 1.0, [{"text": "不能不去", "font": SANS_M, "size": 52, "color": RED}])
    tb(s, 0.85, 6.35, 8, 0.35, [{"text": "蒙想讯息  ·  2026年9月22日", "font": SANS, "size": 14, "color": ON_DARK_DIM}])
    page(s, 15, total, dark=True)
    notes(s, "不再加新论点。四句读完，停在不能不去。")

    return prs


if __name__ == "__main__":
    out = "/workspace/艾多美何以成为100兆韩元企业.pptx"
    build().save(out)
    print(out)
