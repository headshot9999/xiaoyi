#!/usr/bin/env python3
"""艾多美一百兆论证 · 演示文稿"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SERIF = "Noto Serif CJK SC"
SERIF_M = "Noto Serif CJK SC Medium"
SANS = "Noto Sans CJK SC"
SANS_M = "Noto Sans CJK SC Medium"

INK = "1C1916"
INK2 = "3A342F"
PAPER = "F4F0E8"
CREAM = "F7F3EB"
MUTED = "7C756E"
FAINT = "C9C1B6"
LACQUER = "8E2E2A"
LACQUER_DEEP = "6F2421"

W = 13.333333
H = 7.5
MX = 0.78


def rgb(h):
    return RGBColor.from_string(h)


def _anchor(tf, valign):
    body = tf._txBody
    bodyPr = body.find(qn("a:bodyPr"))
    bodyPr.set("anchor", {"top": "t", "middle": "ctr", "bottom": "b"}[valign])
    for key in ("lIns", "rIns", "tIns", "bIns"):
        bodyPr.set(key, "0")


def _run(p, text, font, size, color, tracking=0):
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
    if tracking:
        rPr.set("spc", str(int(tracking * 100)))
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
                _run(
                    p,
                    r["text"],
                    r.get("font", SANS),
                    r["size"],
                    r["color"],
                    r.get("tracking", 0),
                )
        else:
            _run(
                p,
                info["text"],
                info.get("font", SANS),
                info["size"],
                info["color"],
                info.get("tracking", 0),
            )
    return shape


def rect(slide, x, y, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = rgb(fill)
    sh.line.fill.background()
    return sh


def rule(slide, x, y, w, color, pt=1.15):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Pt(pt))
    sh.fill.solid()
    sh.fill.fore_color.rgb = rgb(color)
    sh.line.fill.background()
    return sh


def vrule(slide, x, y, h, color, pt=1.0):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Pt(pt), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = rgb(color)
    sh.line.fill.background()
    return sh


def bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def footer(slide, n, total, dark=False):
    color = "A3988C" if dark else MUTED
    tb(
        slide,
        MX,
        7.05,
        6.5,
        0.28,
        [{"text": "艾多美  ·  一百兆韩元", "font": SANS, "size": 11, "color": color, "tracking": 1.2}],
    )
    tb(
        slide,
        10.4,
        7.05,
        2.15,
        0.28,
        [
            {
                "text": f"{n:02d}  /  {total:02d}",
                "font": SANS,
                "size": 11,
                "color": color,
                "align": "right",
                "tracking": 1.4,
            }
        ],
    )


def kicker(slide, text, dark=False):
    color = "E4C7B4" if dark else LACQUER
    tb(
        slide,
        MX,
        0.42,
        10,
        0.32,
        [{"text": text, "font": SANS_M, "size": 12, "color": color, "tracking": 2.4}],
    )
    rule(slide, MX, 0.82, 0.46, color, 1.35)


def title(slide, text, y=0.98, size=32, color=INK, h=0.62):
    tb(
        slide,
        MX,
        y,
        11.7,
        h,
        [{"text": text, "font": SERIF_M, "size": size, "color": color}],
    )


def new(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def build():
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    prs.core_properties.title = "艾多美何以成为100兆韩元企业"
    prs.core_properties.subject = "蒙想讯息 · 二〇二六年九月二十二日"
    prs.core_properties.author = "艾多美课题论证"
    total = 16

    # 01 cover
    s = new(prs)
    bg(s, INK)
    tb(
        s,
        MX,
        0.62,
        8,
        0.32,
        [{"text": "蒙想讯息   ·   二〇二六年九月二十二日", "font": SANS, "size": 13, "color": "C8B8A4", "tracking": 1.6}],
    )
    rule(s, MX, 1.12, 1.15, LACQUER, 1.5)
    tb(
        s,
        MX,
        2.05,
        11,
        1.15,
        [{"text": "艾多美何以成为", "font": SERIF, "size": 54, "color": CREAM}],
    )
    tb(
        s,
        MX,
        3.25,
        11.5,
        1.2,
        [{"text": "一百兆韩元企业", "font": SERIF, "size": 54, "color": CREAM}],
    )
    tb(
        s,
        MX,
        4.85,
        10,
        0.4,
        [{"text": "把口号写成结构", "font": SERIF_M, "size": 20, "color": "E4C7B4"}],
    )
    tb(
        s,
        MX,
        6.45,
        10,
        0.32,
        [{"text": "一贯工程     一能一品     三 A", "font": SANS, "size": 13, "color": "A3988C", "tracking": 2.2}],
    )
    footer(s, 1, total, dark=True)
    notes(s, "开场不要解释标题。停一拍，直接进入下一页的两个数字。")

    # 02 tenfold
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "问题")
    title(s, "十倍，就在这段年月里")
    tb(s, MX, 2.05, 5, 0.3, [{"text": "眼下的第一", "font": SANS, "size": 14, "color": MUTED, "tracking": 1.5}])
    tb(s, MX, 2.4, 5.4, 1.35, [{"text": "十兆", "font": SERIF, "size": 72, "color": INK}])
    tb(
        s,
        MX,
        3.9,
        5.2,
        0.85,
        [
            {"text": "大约六十年", "font": SANS_M, "size": 18, "color": INK2, "after": 4},
            {"text": "安利。行业里走在前面的那个数字。", "font": SANS, "size": 15, "color": MUTED, "line": 1.35},
        ],
    )
    vrule(s, 6.55, 2.15, 3.15, FAINT, 1.0)
    tb(s, 7.05, 2.05, 5.4, 0.3, [{"text": "要去的地方", "font": SANS, "size": 14, "color": LACQUER, "tracking": 1.5}])
    tb(s, 7.05, 2.4, 5.5, 1.35, [{"text": "一百兆", "font": SERIF, "size": 72, "color": LACQUER}])
    tb(
        s,
        7.05,
        3.9,
        5.3,
        0.9,
        [
            {"text": "大约十五年", "font": SANS_M, "size": 18, "color": INK2, "after": 4},
            {"text": "蒙想还在的年月。不是一百年之后。", "font": SANS, "size": 15, "color": MUTED, "line": 1.35},
        ],
    )
    rule(s, MX, 5.55, 11.75, FAINT, 1.0)
    tb(
        s,
        MX,
        5.75,
        11.7,
        0.7,
        [
            {
                "text": "再往后，是千兆。要比过的，是沃尔玛，是亚马逊。",
                "font": SERIF_M,
                "size": 18,
                "color": INK,
            }
        ],
    )
    footer(s, 2, total)
    notes(s, "先把落差放上桌。十兆用了六十年。一百兆要在大约十五年里走到。听众若不信，后面的每一页都是写给这个不信的。")

    # 03 standard
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "标准")
    title(s, "信念握得住口号")
    tb(
        s,
        MX,
        1.85,
        11.5,
        0.45,
        [{"text": "一百兆若只是被喊出来，它还停在念头里。", "font": SANS, "size": 16, "color": INK2}],
    )
    rule(s, MX, 2.55, 5.15, FAINT, 1.0)
    rule(s, 7.15, 2.55, 5.35, LACQUER, 1.15)
    tb(s, MX, 2.8, 5.3, 0.45, [{"text": "信念", "font": SERIF_M, "size": 28, "color": INK}])
    tb(
        s,
        MX,
        3.5,
        5.2,
        2.2,
        [
            {
                "text": "中心是自己。\n喊一喊，念头会更紧。\n一百兆不会因此发生。",
                "font": SANS,
                "size": 18,
                "color": INK2,
                "line": 1.55,
            }
        ],
    )
    tb(s, 7.15, 2.8, 5.4, 0.45, [{"text": "结构", "font": SERIF_M, "size": 28, "color": LACQUER}])
    tb(
        s,
        7.15,
        3.5,
        5.3,
        2.3,
        [
            {
                "text": "中心是事情本身。\n信，或者不信，它都往那里去。\n论证要写给不信的人。",
                "font": SANS,
                "size": 18,
                "color": INK2,
                "line": 1.55,
            }
        ],
    )
    footer(s, 3, total)
    notes(s, "信念是把自己的想法握紧。要的是另一件事：结构如此。后面所有页，都按“不信的人也能跟上”来讲。")

    # 04 battlefield
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "战场")
    title(s, "先把人分开")
    tb(
        s,
        MX,
        1.82,
        11.6,
        0.55,
        [
            {
                "text": "亚马逊、沃尔玛、Coupang 把场子铺开，说你们自己挑。教科书接着说，高低都做。",
                "font": SANS,
                "size": 16,
                "color": INK2,
            }
        ],
    )
    rows = [
        ("01", "丢掉", "最上约 1%", "嫌便宜。东西再好，也不用。"),
        ("02", "丢掉", "只认便宜", "韩国大约三成。购买力更弱的地方，可以到七成。"),
        ("03", "留下", "绝对品质，绝对价格", "这一层拿下一半，千兆也撑得住。"),
    ]
    y = 2.6
    for idx, (num, verb, head, body) in enumerate(rows):
        if idx:
            rule(s, MX, y, 11.75, FAINT, 1.0)
            y += 0.22
        tb(s, MX, y, 0.7, 0.42, [{"text": num, "font": SERIF, "size": 18, "color": LACQUER}])
        tb(s, 1.6, y, 1.3, 0.42, [{"text": verb, "font": SANS, "size": 16, "color": MUTED}])
        tb(s, 3.05, y, 4.3, 0.42, [{"text": head, "font": SERIF_M, "size": 20, "color": INK}])
        tb(s, 7.5, y, 5.0, 0.5, [{"text": body, "font": SANS, "size": 15, "color": INK2}])
        y += 0.95
    footer(s, 4, total)
    notes(s, "不要全部市场。最上大约百分之一，和只求最贱的那一层，都放下。一百兆的分母，是同时要绝对品质和绝对价格的人，再取一半。")

    # 05 35%
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "津贴")
    title(s, "三十五个点，不能向顾客要")
    tb(s, MX, 1.9, 5.2, 1.5, [{"text": "35%", "font": SERIF, "size": 84, "color": LACQUER}])
    tb(
        s,
        MX,
        3.55,
        4.8,
        1.15,
        [
            {
                "text": "会员侧的津贴。\n早年靠“少了中间商”。\n现在的渠道，大多只过一手。",
                "font": SANS,
                "size": 15,
                "color": INK2,
                "line": 1.4,
            }
        ],
    )
    vrule(s, 6.15, 1.95, 4.35, FAINT, 1.0)
    comps = [
        ("好市多", "毛利为零。年费约八万韩元。买一千万与买一亿，都是这一笔。生产者不另加负担。"),
        ("易买得", "进货比重常在六成到七成。"),
        ("电视购物", "播出费用大约四成二，生产者留下五成八。"),
    ]
    y = 1.9
    for i, (name, desc) in enumerate(comps):
        tb(s, 6.55, y, 5.9, 0.32, [{"text": name, "font": SERIF_M, "size": 16, "color": INK}])
        tb(s, 6.55, y + 0.34, 5.9, 0.7, [{"text": desc, "font": SANS, "size": 13, "color": INK2, "line": 1.25}])
        y += 1.15
        if i < 2:
            rule(s, 6.55, y - 0.12, 5.85, FAINT, 1.0)
    footer(s, 5, total)
    notes(s, "向消费者加价，等于在好市多面前认输。对方降不下来的成本，渠道也降不下来。水要喝，先问水源。")

    # 06 source
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "水源")
    title(s, "工序重做之后，三件事同时成立")
    items = [
        ("会员", "大约半价"),
        ("事业者", "津贴还在"),
        ("公司", "仍有余量"),
    ]
    x = MX
    for i, (who, what) in enumerate(items):
        tb(s, x, 2.15, 3.5, 0.35, [{"text": f"0{i+1}", "font": SERIF, "size": 14, "color": LACQUER, "tracking": 1}])
        tb(s, x, 2.55, 3.5, 0.45, [{"text": who, "font": SANS, "size": 16, "color": MUTED}])
        tb(s, x, 3.05, 3.6, 0.8, [{"text": what, "font": SERIF_M, "size": 32, "color": INK}])
        x += 4.0
    rule(s, MX, 4.35, 11.75, FAINT, 1.0)
    tb(
        s,
        MX,
        4.65,
        11.6,
        1.5,
        [
            {
                "text": "这三十五个点的水源，在生产工序里。\n不是加在价格上，也不是从合作方的利润里硬削下来。",
                "font": SANS,
                "size": 18,
                "color": INK2,
                "line": 1.55,
            }
        ],
    )
    footer(s, 6, total)
    notes(s, "半价、津贴、公司余量，三件同时成立，靠的是同一处水源。下一页用牙刷把这个算法拆开。")

    # 07 toothbrush
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "牙刷")
    title(s, "差出来的，不是毛利")
    tb(s, MX, 1.85, 6, 1.25, [{"text": "1,200", "font": SERIF, "size": 68, "color": LACQUER}])
    tb(s, MX, 3.2, 5.5, 0.7, [{"text": "韩元。二十年，没有人跟上。", "font": SANS_M, "size": 16, "color": INK2}])
    tb(s, 7.15, 2.05, 5, 0.3, [{"text": "市价", "font": SANS, "size": 14, "color": MUTED, "tracking": 1.2}])
    tb(s, 7.15, 2.4, 5.3, 0.9, [{"text": "2,000–3,000", "font": SERIF, "size": 36, "color": INK}])
    # bars
    rect(s, 7.15, 3.45, 5.15, 0.08, FAINT)
    rect(s, 7.15, 3.45, 2.06, 0.08, LACQUER)
    tb(s, 7.15, 3.6, 5.2, 0.3, [{"text": "一千八百韩元，叠着三层", "font": SANS, "size": 13, "color": MUTED}])
    layers = [
        ("制造", "型号一多，每样只做一点"),
        ("信任", "广告，和店铺"),
        ("选择", "挑选，买错，退货"),
    ]
    y = 4.25
    for name, desc in layers:
        tb(s, MX, y, 1.3, 0.4, [{"text": name, "font": SERIF_M, "size": 16, "color": LACQUER}])
        tb(s, 2.2, y, 9.5, 0.4, [{"text": desc, "font": SANS, "size": 16, "color": INK2}])
        y += 0.55
    footer(s, 7, total)
    notes(s, "一个造型。嘴不变，模具不改。腔数做大，原料买吨袋，货款付现金。策展若停在挑一个现成的好货，Coupang 明天就能做。")

    # 08 HemoHIM
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "HemoHIM")
    title(s, "草根煮一煮，为什么那么贵")
    tb(s, MX, 1.9, 5.2, 1.05, [{"text": "770,000", "font": SERIF, "size": 48, "color": MUTED}])
    tb(s, 6.35, 2.15, 0.8, 0.55, [{"text": "→", "font": SERIF, "size": 32, "color": LACQUER, "align": "center"}])
    tb(s, 7.2, 1.85, 5.3, 1.15, [{"text": "76,500", "font": SERIF, "size": 54, "color": LACQUER}])
    tb(
        s,
        MX,
        3.15,
        11.5,
        0.4,
        [{"text": "价格停在十年前。品质走到十年后。依据是一个月十万盒。", "font": SANS, "size": 16, "color": INK2}],
    )
    facts = [
        ("0.5  →  18 吨", "三十六倍。操作的人仍是一个。"),
        ("干燥与切断", "省掉。鲜品洗净，直接取汁。"),
        ("合同栽培", "药材没有弹性。锁在涨价之前。"),
    ]
    y = 3.85
    for head, body in facts:
        rule(s, MX, y, 11.75, FAINT, 1.0)
        tb(s, MX, y + 0.16, 4.3, 0.45, [{"text": head, "font": SERIF_M, "size": 18, "color": INK}])
        tb(s, 5.3, y + 0.18, 7.1, 0.45, [{"text": body, "font": SANS, "size": 16, "color": INK2}])
        y += 0.78
    footer(s, 8, total)
    notes(s, "罐子从零点五吨到十八吨。药房要烘干、切断；大罐用鲜品，这两道就没有了。十月起按七万六千五百韩元卖。")

    # 09 integrated process
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "一贯工程")
    title(s, "重做的是整条线")
    cols = [
        ("01", "走到原料", "调配不是源头。温度、压力、配比，向原料厂要。产线怎么转，图纸在设备厂手里。"),
        ("02", "先锁住，再放量", "工业品可以连夜加产。药材要一两年。市场做大之前，工序先放到不能再省。"),
        ("03", "会做，才交出运转", "设备我们出，原料我们买，对方只运转。一起，可以把小公司做到一百亿。不一起，自己做。"),
    ]
    x = MX
    for i, (num, head, body) in enumerate(cols):
        if i:
            vrule(s, x - 0.28, 2.05, 3.7, FAINT, 1.0)
        tb(s, x, 2.05, 3.4, 0.35, [{"text": num, "font": SERIF, "size": 14, "color": LACQUER}])
        tb(s, x, 2.5, 3.45, 0.9, [{"text": head, "font": SERIF_M, "size": 22, "color": INK}])
        tb(s, x, 3.55, 3.4, 2.2, [{"text": body, "font": SANS, "size": 15, "color": INK2, "line": 1.45}])
        x += 3.95
    footer(s, 9, total)
    notes(s, "绝对价格的定义：谁来做，都做不到比这更便宜。一百兆不是“等对方肯配合”。跟，就一起；不跟，优化也不停。")

    # 10 daiso
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "对手")
    title(s, "生产侧，对方已经走到了")
    tb(s, MX, 2.0, 5.2, 0.32, [{"text": "大创的销售", "font": SANS, "size": 14, "color": MUTED, "tracking": 1.2}])
    tb(s, MX, 2.4, 5.5, 1.15, [{"text": "4.5 兆", "font": SERIF, "size": 60, "color": INK}])
    tb(s, MX, 3.7, 5.2, 0.7, [{"text": "增长维持在三成以上。\n策展、放量、进产线，它都做。", "font": SANS, "size": 15, "color": INK2, "line": 1.4}])
    vrule(s, 6.55, 2.1, 2.7, FAINT, 1.0)
    tb(s, 7.05, 2.0, 5.2, 0.32, [{"text": "艾多美眼下的增长", "font": SANS, "size": 14, "color": LACQUER, "tracking": 1.2}])
    tb(s, 7.05, 2.4, 5.4, 1.15, [{"text": "约 3%", "font": SERIF, "size": 60, "color": LACQUER}])
    tb(s, 7.05, 3.7, 5.2, 0.7, [{"text": "品质更好、设计更好，\n解释不了这一仗。", "font": SANS, "size": 15, "color": INK2, "line": 1.4}])
    rule(s, MX, 5.05, 11.75, FAINT, 1.0)
    tb(
        s,
        MX,
        5.3,
        11.6,
        1.1,
        [
            {
                "text": "便宜不等于粗糙。中国制造大约三成，其余来自韩国和其他地方。\n生产上能挖的，对方大多已经挖到了。差别要到销售这一侧去找。",
                "font": SANS,
                "size": 16,
                "color": INK2,
                "line": 1.45,
            }
        ],
    )
    footer(s, 10, total)
    notes(s, "大创比卖场更要认真看。经营者从制造里出来，下一代已经进入化妆品。说品牌更高，是不够的。")

    # 11 organization
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "组织")
    title(s, "店不存在，队伍先走")
    tb(s, MX, 1.9, 5.4, 0.4, [{"text": "大创", "font": SERIF_M, "size": 22, "color": INK}])
    tb(
        s,
        MX,
        2.5,
        5.3,
        1.5,
        [
            {
                "text": "必须先有店，人再来。\n没有店，连找上门的路都没有。\n店的成本，和到店的成本，都是实的。",
                "font": SANS,
                "size": 16,
                "color": INK2,
                "line": 1.45,
            }
        ],
    )
    tb(s, 7.15, 1.9, 5.4, 0.4, [{"text": "艾多美", "font": SERIF_M, "size": 22, "color": LACQUER}])
    tb(
        s,
        7.15,
        2.5,
        5.3,
        1.6,
        [
            {
                "text": "没有店。人自己用，也介绍。\n半价的体验代替广告。\n省下来的，就是那三十五个点。",
                "font": SANS,
                "size": 16,
                "color": INK2,
                "line": 1.45,
            }
        ],
    )
    # flywheel
    rule(s, MX, 4.45, 11.75, FAINT, 1.0)
    steps = ["绝对价格", "会员", "产量", "成本再降"]
    x = MX
    for i, step in enumerate(steps):
        tb(s, x, 4.7, 2.15, 0.45, [{"text": step, "font": SERIF_M, "size": 16, "color": INK}])
        if i < 3:
            tb(s, x + 1.85, 4.68, 0.4, 0.4, [{"text": "—", "font": SERIF, "size": 16, "color": LACQUER, "align": "center"}])
        x += 2.95
    tb(
        s,
        MX,
        5.4,
        11.6,
        1.15,
        [
            {
                "text": "摩洛哥，会员十万以上。孟加拉，约二十万。收入差这么远，卖的是同一套产品。\n轮子没有刹车。邻里若觉得街角有差不多的东西、价格只有一半的一半，话就传不出去。",
                "font": SANS,
                "size": 15,
                "color": INK2,
                "line": 1.45,
            }
        ],
    )
    footer(s, 11, total)
    notes(s, "组织先于店铺移动。非洲、孟加拉、秘鲁、巴拿马，都是队伍先走。绝对价格招来会员，会员堆出产量，产量把成本再压低。")

    # 12 curation
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "策展")
    title(s, "不停在货架上")
    qs = [
        ("留下什么", "一个品类，一个。品质不拿来换价格。顾客不必在几百个相近的东西之间比较。"),
        ("如何守住", "从原物到流通。工厂不必座座自有。原料、设备基准和数据，要在自己手里。循环的第一步，是对方看得见的愿景。"),
        ("对准谁", "同一种产品，不再对所有人说同一句话。身体不同，建议就不同。"),
    ]
    y = 1.9
    for head, body in qs:
        tb(s, MX, y, 3.3, 0.85, [{"text": head, "font": SERIF_M, "size": 22, "color": INK}])
        tb(s, 4.3, y, 8.2, 0.95, [{"text": body, "font": SANS, "size": 16, "color": INK2, "line": 1.35}])
        y += 1.25
        if y < 5.5:
            rule(s, MX, y - 0.22, 11.75, FAINT, 1.0)
    tb(
        s,
        MX,
        5.85,
        11.6,
        0.7,
        [
            {
                "text": "五百个，不能按五百个算。亚马逊一把牙刷就有几百款。我们一把，顶被挑剩的那一个。",
                "font": SANS_M,
                "size": 15,
                "color": INK,
            }
        ],
    )
    footer(s, 12, total)
    notes(s, "信息和商品溢出来的时候，人要的不是更多，是谁按什么标准替他选。大量生产人人懂，进不去，是因为两边都在躲风险。愿景要先被看见。")

    # 13 3A
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "三 A")
    title(s, "三台引擎，一个结构")
    engines = [
        ("APP", "个人平台", "自己的店，自己的健康，自己的教育，收成一处。一小时，变成五分钟。平台本身就是人工智能。"),
        ("A-Care", "艾护理", "从“好东西请用”，进到血、饮食，和以后的十年。越用越准。人因此留下来。"),
        ("AZA", "阿扎", "鸡蛋、油、加油，接进同一处。十万个品目负责人每天走进来。一百兆，由大约五百个主力品目赚。"),
    ]
    x = MX
    for i, (en, cn, body) in enumerate(engines):
        if i:
            vrule(s, x - 0.28, 1.95, 3.55, FAINT, 1.0)
        tb(s, x, 1.95, 3.45, 0.4, [{"text": en, "font": SERIF, "size": 14, "color": LACQUER, "tracking": 1.2}])
        tb(s, x, 2.4, 3.45, 0.55, [{"text": cn, "font": SERIF_M, "size": 26, "color": INK}])
        tb(s, x, 3.2, 3.4, 2.2, [{"text": body, "font": SANS, "size": 15, "color": INK2, "line": 1.45}])
        x += 3.95
    rule(s, MX, 5.7, 11.75, FAINT, 1.0)
    tb(
        s,
        MX,
        5.9,
        11.6,
        0.7,
        [
            {
                "text": "事业者从推销变成顾问，变成这家平台的主人。消费、推荐和组织，变成留得下的资产。",
                "font": SANS,
                "size": 16,
                "color": INK2,
            }
        ],
    )
    footer(s, 13, total)
    notes(s, "三A不是三款应用。艾护理负责关系，阿扎负责每天来，个人平台负责把事业做成方法。加油走联名卡的积分，不把加油券当商品卖。")

    # 14 four
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "结论")
    title(s, "四件事同时为真")
    lines = [
        ("01", "这条链不容易被照抄", "抄一张价格表容易。抄育苗、栽培、图纸、专用线和数据，要同时有愿景和敢先投入的人。"),
        ("02", "不附带对方肯配合", "跟，就一起把小公司做大。不跟，优化也不停。"),
        ("03", "一半，就是千兆的底", "不要全部市场。绝对品质、绝对价格这一层的一半，已经够。"),
        ("04", "人留下，产量才留下", "健康越用越准，日常开支接在一处，内容和组织长在自己的平台上。"),
    ]
    y = 1.85
    for num, head, body in lines:
        tb(s, MX, y, 0.7, 0.4, [{"text": num, "font": SERIF, "size": 16, "color": LACQUER}])
        tb(s, 1.65, y - 0.02, 4.5, 0.4, [{"text": head, "font": SERIF_M, "size": 18, "color": INK}])
        tb(s, 6.3, y, 6.2, 0.7, [{"text": body, "font": SANS, "size": 14, "color": INK2, "line": 1.25}])
        y += 1.15
    footer(s, 14, total)
    notes(s, "一百兆是这四句同时转起来的结果。可以分开背，不能分开成立。")

    # 15 two places
    s = new(prs)
    bg(s, PAPER)
    kicker(s, "落地")
    title(s, "同一套算法")
    tb(s, MX, 2.05, 5.4, 0.45, [{"text": "台湾", "font": SERIF_M, "size": 32, "color": INK}])
    tb(
        s,
        MX,
        2.75,
        5.35,
        2.6,
        [
            {
                "text": "这个岛靠把成本和品质管到根上过活。\n在工厂里拆过成本的人，读成分，算价格，相信用过的人。\n绝对品质、绝对价格，在这里传得快。\n站住了，华文圈就有了标准。",
                "font": SANS,
                "size": 16,
                "color": INK2,
                "line": 1.5,
            }
        ],
    )
    vrule(s, 6.6, 2.1, 3.5, FAINT, 1.0)
    tb(s, 7.1, 2.05, 5.4, 0.45, [{"text": "澳洲", "font": SERIF_M, "size": 32, "color": INK}])
    tb(
        s,
        7.1,
        2.75,
        5.3,
        2.6,
        [
            {
                "text": "麦卢卡的价格写在熟成里，\n不写在分装上。\n只谈成品价，结构不动。\n先和蜂农锁量，再谈这一瓶多少钱。",
                "font": SANS,
                "size": 16,
                "color": INK2,
                "line": 1.5,
            }
        ],
    )
    footer(s, 15, total)
    notes(s, "两地一南一北。台湾说明认得成本的市场为什么吃这套逻辑。澳洲说明一贯工程怎样落在一个具体的产地。")

    # 16 close
    s = new(prs)
    bg(s, INK)
    tb(
        s,
        MX,
        0.58,
        8,
        0.3,
        [{"text": "所以", "font": SANS_M, "size": 13, "color": "E4C7B4", "tracking": 3}],
    )
    rule(s, MX, 1.05, 0.46, LACQUER, 1.5)
    lines = ["生产被重做", "消费者被组织", "选择被拿掉", "三十五个点有了出处"]
    y = 1.4
    for line in lines:
        tb(s, MX, y, 11, 0.55, [{"text": line, "font": SERIF, "size": 28, "color": CREAM}])
        y += 0.72
    tb(s, MX, 4.55, 11, 1.0, [{"text": "不能不去", "font": SERIF, "size": 54, "color": "E4C7B4"}])
    tb(
        s,
        MX,
        6.35,
        10,
        0.35,
        [{"text": "蒙想讯息    ·    二〇二六年九月二十二日", "font": SANS, "size": 13, "color": "A3988C", "tracking": 1.2}],
    )
    footer(s, 16, total, dark=True)
    notes(s, "收束不再加新论点。四句读完，停在“不能不去”。问的时候，回到水源、组织、三A，这三处。")

    return prs


if __name__ == "__main__":
    prs = build()
    out = "/workspace/艾多美何以成为100兆韩元企业.pptx"
    prs.save(out)
    print(out)
