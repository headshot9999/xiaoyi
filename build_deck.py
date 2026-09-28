#!/usr/bin/env python3
"""艾多美一百兆论证 · 演示文稿"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SANS = "Noto Sans CJK SC"
MED = "Noto Sans CJK SC Medium"

WHITE = "FFFFFF"
INK = "111111"
BODY = "3D3D3D"
MUTED = "8A8A8A"
HAIR = "E4E4E4"
RED = "C8102E"
DARK = "0E0E0E"
DIM = "9A9A9A"
PAPER = "F6F6F6"

W = 13.333333
H = 7.5
ML = 0.78


def rgb(value):
    return RGBColor.from_string(value)


def _anchor(tf, valign):
    body = tf._txBody.find(qn("a:bodyPr"))
    body.set("anchor", {"top": "t", "middle": "ctr", "bottom": "b"}[valign])
    for key in ("lIns", "rIns", "tIns", "bIns"):
        body.set(key, "0")


def _run(p, text, font, size, color):
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = False
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
                _run(p, item["text"], item.get("font", SANS), item["size"], item["color"])
        else:
            _run(p, info["text"], info.get("font", SANS), info["size"], info["color"])
    return shape


def rect(slide, x, y, w, h, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.fill.background()
    return shape


def bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def new(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def folio(slide, n, dark=False):
    tb(
        slide, 11.4, 7.08, 1.35, 0.24,
        [{"text": f"{n:02d}", "font": SANS, "size": 12, "color": "5C5C5C" if dark else "C4C4C4", "align": "right"}],
    )


def title(slide, text, y=0.42):
    tb(slide, ML, y, 11.6, 0.62, [{"text": text, "font": MED, "size": 34, "color": INK}])


def build():
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    prs.core_properties.title = "艾多美何以成为100兆韩元企业"
    prs.core_properties.subject = "蒙想讯息 · 2026年9月22日"

    # 01
    s = new(prs)
    bg(s, DARK)
    rect(s, 0, 0, 0.1, H, RED)
    tb(s, 0.9, 1.55, 10, 0.32, [{"text": "蒙想讯息  ·  2026年9月22日", "font": SANS, "size": 15, "color": DIM}])
    tb(
        s, 0.88, 2.35, 11.6, 2.3,
        [
            {"text": "艾多美何以成为", "font": MED, "size": 52, "color": WHITE, "after": 2},
            {"text": "一百兆韩元企业", "font": MED, "size": 52, "color": WHITE},
        ],
    )
    rect(s, 0.92, 5.0, 1.15, 0.04, RED)
    tb(s, 0.92, 5.25, 8, 0.4, [{"text": "把口号写成结构", "font": SANS, "size": 20, "color": "CFCFCF"}])
    folio(s, 1, dark=True)
    notes(s, "停一拍。下一页只给两个数字。")

    # 02
    s = new(prs)
    bg(s, WHITE)
    rect(s, 0, 0, 6.45, H, DARK)
    tb(s, 0.78, 1.85, 5, 0.3, [{"text": "眼下的第一", "font": SANS, "size": 15, "color": DIM}])
    tb(s, 0.72, 2.25, 5.4, 1.35, [{"text": "十兆", "font": MED, "size": 84, "color": WHITE}])
    tb(s, 0.78, 3.75, 5, 0.42, [{"text": "大约六十年", "font": MED, "size": 22, "color": WHITE}])
    tb(s, 0.78, 4.35, 5, 0.7, [{"text": "安利。行业里走在前面的数字。", "font": SANS, "size": 16, "color": DIM}])
    tb(s, 7.05, 1.85, 5.5, 0.3, [{"text": "要去的地方", "font": SANS, "size": 15, "color": RED}])
    tb(s, 6.98, 2.25, 5.8, 1.35, [{"text": "一百兆", "font": MED, "size": 84, "color": RED}])
    tb(s, 7.05, 3.75, 5.5, 0.42, [{"text": "大约十五年", "font": MED, "size": 22, "color": INK}])
    tb(s, 7.05, 4.35, 5.5, 0.7, [{"text": "蒙想还在的年月。", "font": SANS, "size": 16, "color": BODY}])
    rect(s, 7.05, 5.55, 5.4, 0.015, HAIR)
    tb(s, 7.05, 5.8, 5.5, 0.9, [{"text": "再往后是千兆。\n要比过的，是沃尔玛，是亚马逊。", "font": SANS, "size": 16, "color": BODY, "line": 1.5}])
    folio(s, 2)
    notes(s, "十兆用了大约六十年。一百兆要在大约十五年里走到。后面都是写给不信的人。")

    # 03
    s = new(prs)
    bg(s, WHITE)
    title(s, "信念握得住口号")
    tb(s, ML, 1.7, 5.2, 0.5, [{"text": "信念", "font": MED, "size": 28, "color": INK}])
    tb(
        s, ML, 2.45, 5.2, 2.4,
        [{"text": "中心是自己。\n喊一喊，念头会更紧。\n一百兆不会因此发生。", "font": SANS, "size": 20, "color": BODY, "line": 1.6}],
    )
    rect(s, 6.85, 1.45, 6.49, 5.35, DARK)
    tb(s, 7.3, 2.15, 5.4, 0.55, [{"text": "结构", "font": MED, "size": 28, "color": WHITE}])
    tb(
        s, 7.3, 2.95, 5.4, 2.6,
        [{"text": "中心是事情本身。\n信，或者不信，它都往那里去。\n论证要写给不信的人。", "font": SANS, "size": 20, "color": "E8E8E8", "line": 1.6}],
    )
    folio(s, 3)
    notes(s, "信念是把自己的想法握紧。要交的是结构。")

    # 04
    s = new(prs)
    bg(s, WHITE)
    title(s, "先把人分开")
    tb(s, ML, 1.25, 11.5, 0.4, [{"text": "场子铺开，是让人自己挑。高低都做，方向就会松。", "font": SANS, "size": 16, "color": MUTED}])
    rows = [
        ("01", "丢掉最上约 1%", "嫌便宜。东西再好，也不用。", False),
        ("02", "丢掉只认便宜的人", "韩国大约三成。购买力更弱的地方，可以到七成。", False),
        ("03", "留下绝对品质，绝对价格", "这一层拿下一半，千兆也撑得住。", True),
    ]
    y = 1.95
    for num, head, body, hot in rows:
        color = RED if hot else INK
        tb(s, ML, y, 1.15, 0.5, [{"text": num, "font": MED, "size": 20, "color": RED if hot else MUTED}])
        tb(s, 2.15, y - 0.04, 9.5, 0.5, [{"text": head, "font": MED, "size": 26, "color": color}])
        tb(s, 2.15, y + 0.58, 9.5, 0.4, [{"text": body, "font": SANS, "size": 16, "color": BODY}])
        y += 1.55
    folio(s, 4)
    notes(s, "不要全部市场。一百兆的分母，是同时要绝对品质和绝对价格的人，再取一半。")

    # 05
    s = new(prs)
    bg(s, WHITE)
    title(s, "不能向顾客要")
    tb(s, ML, 1.45, 5.4, 1.45, [{"text": "35%", "font": MED, "size": 92, "color": RED}])
    tb(s, ML, 3.15, 5.2, 1.5, [{"text": "会员侧的津贴。\n渠道现在大多只过一手，\n不能再靠少了中间商来解释。", "font": SANS, "size": 16, "color": BODY, "line": 1.5}])
    items = [
        ("好市多", "毛利为零。年费约八万韩元。\n买一千万与买一亿，都是这一笔。"),
        ("易买得", "进货比重常在六成到七成。"),
        ("电视购物", "播出费用大约四成二。"),
    ]
    y = 1.55
    for i, (name, desc) in enumerate(items):
        tb(s, 6.7, y, 5.8, 0.4, [{"text": name, "font": MED, "size": 20, "color": INK}])
        tb(s, 6.7, y + 0.46, 5.8, 0.75, [{"text": desc, "font": SANS, "size": 15, "color": BODY, "line": 1.35}])
        y += 1.5
        if i < 2:
            rect(s, 6.7, y - 0.18, 5.7, 0.012, HAIR)
    folio(s, 5)
    notes(s, "向消费者加价，等于在好市多面前认输。水源要到工序里去找。")

    # 06
    s = new(prs)
    bg(s, WHITE)
    title(s, "三件事同时成立")
    cols = [("01", "会员", "大约半价"), ("02", "事业者", "津贴还在"), ("03", "公司", "仍有余量")]
    x = ML
    for num, who, what in cols:
        tb(s, x, 2.15, 3.6, 0.35, [{"text": num, "font": MED, "size": 14, "color": RED}])
        tb(s, x, 2.65, 3.6, 0.4, [{"text": who, "font": SANS, "size": 16, "color": MUTED}])
        tb(s, x, 3.2, 3.7, 1.0, [{"text": what, "font": MED, "size": 40, "color": INK}])
        x += 4.05
    rect(s, ML, 4.85, 11.75, 0.012, HAIR)
    tb(
        s, ML, 5.2, 11.2, 1.3,
        [{"text": "三十五个点的水源，在生产工序里。\n不是加在价格上，也不是硬削合作方的利润。", "font": SANS, "size": 20, "color": INK, "line": 1.55}],
    )
    folio(s, 6)
    notes(s, "半价、津贴、公司余量，靠的是同一处水源。")

    # 07
    s = new(prs)
    bg(s, WHITE)
    title(s, "差出来的，不是毛利")
    tb(s, ML, 1.55, 6, 1.25, [{"text": "1,200", "font": MED, "size": 78, "color": RED}])
    tb(s, ML, 2.95, 5, 0.35, [{"text": "韩元  ·  二十年没有人跟上", "font": SANS, "size": 16, "color": BODY}])
    tb(s, 7.15, 1.6, 5.3, 0.3, [{"text": "市价", "font": SANS, "size": 14, "color": MUTED}])
    tb(s, 7.15, 1.95, 5.4, 0.7, [{"text": "2,000 – 3,000", "font": MED, "size": 32, "color": INK}])
    rect(s, 7.15, 2.85, 5.2, 0.14, HAIR)
    rect(s, 7.15, 2.85, 2.08, 0.14, RED)
    layers = [("制造", "型号一多，每样只做一点"), ("信任", "广告，和店铺"), ("选择", "挑选，买错，退货")]
    y = 3.3
    for name, desc in layers:
        tb(s, 7.15, y, 1.2, 0.38, [{"text": name, "font": MED, "size": 16, "color": RED}])
        tb(s, 8.45, y, 4, 0.38, [{"text": desc, "font": SANS, "size": 16, "color": INK}])
        y += 0.55
    rect(s, ML, 5.55, 11.75, 0.012, HAIR)
    tb(s, ML, 5.8, 11.5, 0.85, [{"text": "一个造型。腔数做大，原料买吨袋，货款付现金。\n停在挑一件现成的好货，别人明天就能做。", "font": SANS, "size": 16, "color": BODY, "line": 1.45}])
    folio(s, 7)
    notes(s, "把产量聚到一个造型上，成本才下得去。")

    # 08
    s = new(prs)
    bg(s, WHITE)
    title(s, "为什么能卖到这个价")
    rect(s, 0, 1.4, 6.45, 3.15, PAPER)
    rect(s, 6.45, 1.4, 6.89, 3.15, DARK)
    tb(s, 0.78, 1.65, 5, 0.3, [{"text": "原先", "font": SANS, "size": 14, "color": MUTED}])
    tb(s, 0.74, 2.1, 5.4, 1.05, [{"text": "770,000", "font": MED, "size": 44, "color": INK}])
    tb(s, 0.78, 3.35, 5, 0.35, [{"text": "韩元", "font": SANS, "size": 15, "color": MUTED}])
    tb(s, 7.15, 1.65, 5.4, 0.3, [{"text": "现在", "font": SANS, "size": 14, "color": "F0A0A0"}])
    tb(s, 7.1, 2.05, 5.6, 1.15, [{"text": "76,500", "font": MED, "size": 52, "color": WHITE}])
    tb(s, 7.15, 3.4, 5, 0.35, [{"text": "韩元", "font": SANS, "size": 15, "color": "F0A0A0"}])
    tb(
        s, ML, 4.72, 11.7, 0.38,
        [{"text": "价格停在十年前，品质走到十年后。依据是一个月十万盒。", "font": MED, "size": 16, "color": INK}],
    )
    facts = [
        ("0.5 吨到 18 吨", "三十六倍，仍是一个人。"),
        ("干燥与切断", "省掉。鲜品直接取汁。"),
        ("合同栽培", "锁在涨价之前。"),
    ]
    x = ML
    for i, (head, body) in enumerate(facts):
        if i:
            rect(s, x - 0.22, 5.35, 0.012, 0.95, HAIR)
        tb(s, x, 5.28, 3.5, 0.4, [{"text": head, "font": MED, "size": 16, "color": INK}])
        tb(s, x, 5.72, 3.5, 0.55, [{"text": body, "font": SANS, "size": 15, "color": BODY}])
        x += 4.05
    folio(s, 8)
    notes(s, "价格停在十年前，品质走到十年后。依据是一个月十万盒。")

    # 09
    s = new(prs)
    bg(s, WHITE)
    title(s, "重做的是整条线")
    cols = [
        ("01", "走到原料", "调配不是源头。温度、压力、配比，向原料厂要。图纸在设备厂手里。"),
        ("02", "先锁住，再放量", "工业品可以连夜加产。药材要一两年。市场做大之前，先把工序放到不能再省。"),
        ("03", "会做，才交出运转", "设备我们出，原料我们买。一起，小公司可以做到一百亿。不一起，自己做。"),
    ]
    x = ML
    for num, head, body in cols:
        rect(s, x, 1.6, 0.38, 0.045, RED)
        tb(s, x, 1.85, 3.5, 0.32, [{"text": num, "font": MED, "size": 13, "color": RED}])
        tb(s, x, 2.3, 3.6, 1.05, [{"text": head, "font": MED, "size": 26, "color": INK}])
        tb(s, x, 3.55, 3.55, 2.0, [{"text": body, "font": SANS, "size": 16, "color": BODY, "line": 1.5}])
        x += 4.05
    tb(s, ML, 6.15, 11.5, 0.45, [{"text": "谁来做，都无法更便宜。跟，就一起。不跟，自己做。", "font": MED, "size": 18, "color": INK}])
    folio(s, 9)
    notes(s, "一百兆不附带“对方肯配合”这个条件。")

    # 10
    s = new(prs)
    bg(s, WHITE)
    title(s, "生产侧，对方已经走到了")
    tb(s, ML, 1.7, 5, 0.32, [{"text": "大创的销售", "font": SANS, "size": 15, "color": MUTED}])
    tb(s, ML - 0.04, 2.1, 5.5, 1.2, [{"text": "4.5 兆", "font": MED, "size": 64, "color": INK}])
    tb(s, ML, 3.5, 5.2, 0.9, [{"text": "增长在三成以上。\n策展、放量、进产线，它都做。", "font": SANS, "size": 16, "color": BODY, "line": 1.45}])
    rect(s, 6.55, 1.85, 0.012, 2.7, HAIR)
    tb(s, 7.05, 1.7, 5.4, 0.32, [{"text": "艾多美眼下的增长", "font": SANS, "size": 15, "color": RED}])
    tb(s, 7.0, 2.1, 5.5, 1.2, [{"text": "约 3%", "font": MED, "size": 64, "color": RED}])
    tb(s, 7.05, 3.5, 5.3, 0.9, [{"text": "品质更好、设计更好，\n解释不了这一仗。", "font": SANS, "size": 16, "color": BODY, "line": 1.45}])
    rect(s, ML, 5.15, 11.75, 0.012, HAIR)
    tb(s, ML, 5.45, 11.5, 1.0, [{"text": "便宜不等于粗糙。生产上能挖的，对方大多已经挖到。\n差别在销售这一侧。", "font": SANS, "size": 18, "color": INK, "line": 1.5}])
    folio(s, 10)
    notes(s, "大创比一般卖场更要认真看。品牌更高，不够。")

    # 11
    s = new(prs)
    bg(s, WHITE)
    title(s, "店不存在，队伍先走")
    tb(s, ML, 1.5, 5.3, 0.45, [{"text": "大创", "font": MED, "size": 24, "color": INK}])
    tb(s, ML, 2.15, 5.3, 1.3, [{"text": "必须先有店，人再来。\n店的成本，和到店的成本，都是实的。", "font": SANS, "size": 16, "color": BODY, "line": 1.5}])
    rect(s, 6.7, 1.4, 6.64, 2.55, DARK)
    tb(s, 7.1, 1.65, 5.6, 0.45, [{"text": "艾多美", "font": MED, "size": 24, "color": WHITE}])
    tb(s, 7.1, 2.25, 5.7, 1.3, [{"text": "没有店。人自己用，也介绍。\n半价的体验代替广告。省下的，就是那 35%。", "font": SANS, "size": 16, "color": "E6E6E6", "line": 1.5}])
    tb(
        s, ML, 4.35, 11.7, 0.5,
        [{
            "align": "center",
            "runs": [
                {"text": "绝对价格", "font": MED, "size": 20, "color": RED},
                {"text": "    —    ", "font": MED, "size": 20, "color": RED},
                {"text": "会员", "font": MED, "size": 20, "color": INK},
                {"text": "    —    ", "font": MED, "size": 20, "color": RED},
                {"text": "产量", "font": MED, "size": 20, "color": INK},
                {"text": "    —    ", "font": MED, "size": 20, "color": RED},
                {"text": "成本再降", "font": MED, "size": 20, "color": INK},
            ]
        }],
    )
    tb(
        s, ML, 5.25, 11.6, 1.2,
        [{"text": "摩洛哥会员十万以上，孟加拉约二十万。收入差这么远，卖的是同一套。\n轮子没有刹车。街上若有差不多的东西、价格低很多，话就传不出去。", "font": SANS, "size": 16, "color": BODY, "line": 1.5}],
    )
    folio(s, 11)
    notes(s, "组织先于店铺移动。价格招来会员，会员堆出产量，产量再把成本压低。")

    # 12
    s = new(prs)
    bg(s, WHITE)
    title(s, "不停在货架上")
    qs = [
        ("留下什么", "一个品类，一个。品质不拿来换价格。"),
        ("如何守住", "从原物管到流通。工厂不必座座自有，标准和数据要在自己手里。"),
        ("对准谁", "同一种产品，不再对所有人说同一句话。"),
    ]
    y = 1.6
    for head, body in qs:
        tb(s, ML, y, 3.4, 0.55, [{"text": head, "font": MED, "size": 24, "color": INK}])
        tb(s, 4.5, y + 0.06, 8, 0.7, [{"text": body, "font": SANS, "size": 18, "color": BODY, "line": 1.35}])
        y += 1.15
    tb(s, ML, 5.35, 11.5, 0.9, [{"text": "五百个，不能按五百个算。\n一把牙刷，顶别人那几百款里被挑剩的一个。", "font": MED, "size": 22, "color": RED, "line": 1.4}])
    folio(s, 12)
    notes(s, "大量生产人人懂。进不去，是因为两边都在躲风险。愿景要先被看见。")

    # 13
    s = new(prs)
    bg(s, WHITE)
    title(s, "三台引擎，一个结构")
    engines = [
        ("APP", "个人平台", "店、健康、教育，收成一处。\n一小时，变成五分钟。"),
        ("A-Care", "艾护理", "从“请用”，进到血和饮食，\n以及以后的十年。越用越准。"),
        ("AZA", "阿扎", "鸡蛋、油、加油接进来。\n一百兆，靠大约五百个主力品目。"),
    ]
    x = ML
    for en, cn, body in engines:
        rect(s, x, 1.6, 0.42, 0.045, RED)
        tb(s, x, 1.85, 3.5, 0.32, [{"text": en, "font": MED, "size": 14, "color": RED}])
        tb(s, x, 2.3, 3.6, 0.6, [{"text": cn, "font": MED, "size": 28, "color": INK}])
        tb(s, x, 3.15, 3.55, 1.5, [{"text": body, "font": SANS, "size": 16, "color": BODY, "line": 1.5}])
        x += 4.05
    rect(s, ML, 5.15, 11.75, 0.012, HAIR)
    tb(s, ML, 5.45, 11.5, 1.0, [{"text": "事业者从推销变成顾问，变成平台的主人。\n消费、推荐和组织，变成留得下的资产。", "font": SANS, "size": 18, "color": INK, "line": 1.5}])
    folio(s, 13)
    notes(s, "艾护理负责关系，阿扎负责每天来，个人平台负责把事业做成方法。")

    # 14
    s = new(prs)
    bg(s, WHITE)
    title(s, "四件事同时为真")
    lines = [
        ("01", "这条链不容易被照抄", "价格表好抄。育苗、栽培、图纸和数据，要有人先投入。"),
        ("02", "不附带对方肯配合", "跟，就一起做大。不跟，优化也不停。"),
        ("03", "一半，就是千兆的底", "不要全部市场。这一层人的一半，已经够。"),
        ("04", "人留下，产量才留下", "越用越准，开支接在一处，组织长在自己的平台上。"),
    ]
    y = 1.5
    for num, head, body in lines:
        tb(s, ML, y, 0.9, 0.45, [{"text": num, "font": MED, "size": 18, "color": RED}])
        tb(s, 1.9, y - 0.02, 10.5, 0.45, [{"text": head, "font": MED, "size": 24, "color": INK}])
        tb(s, 1.9, y + 0.5, 10.5, 0.38, [{"text": body, "font": SANS, "size": 15, "color": BODY}])
        y += 1.28
    folio(s, 14)
    notes(s, "四句同时转，才是一百兆。可以分开讲，不能分开成立。")

    # 15
    s = new(prs)
    bg(s, DARK)
    rect(s, 0, 0, 0.1, H, RED)
    lines = ["生产被重做", "消费者被组织", "选择被拿掉", "三十五个点有了出处"]
    y = 0.95
    for line in lines:
        tb(s, 0.9, y, 11, 0.58, [{"text": line, "font": MED, "size": 30, "color": WHITE}])
        y += 0.78
    tb(s, 0.88, 4.35, 11, 1.05, [{"text": "不能不去", "font": MED, "size": 58, "color": RED}])
    tb(s, 0.92, 6.4, 8, 0.32, [{"text": "蒙想讯息  ·  2026年9月22日", "font": SANS, "size": 14, "color": DIM}])
    folio(s, 15, dark=True)
    notes(s, "不再加新论点。停在不能不去。")

    return prs


if __name__ == "__main__":
    path = "/workspace/艾多美何以成为100兆韩元企业.pptx"
    build().save(path)
    print(path)
