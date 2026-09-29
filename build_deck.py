#!/usr/bin/env python3
"""艾多美一百兆论证。每一页是一句可被反驳的判断，下面是证据。"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SANS = "Noto Sans CJK SC"
MED = "Noto Sans CJK SC Medium"

WHITE = "FFFFFF"
INK = "141414"
BODY = "2C2C2C"
MUTED = "6A6A6A"
HAIR = "E2E2E2"
RED = "9E1B1B"
DARK = "121212"
DIM = "9A9A9A"

W = 13.333333
H = 7.5
ML = 0.72


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
        slide, 11.55, 7.1, 1.15, 0.22,
        [{"text": f"{n:02d}", "font": SANS, "size": 11, "color": "6A6A6A" if dark else "C8C8C8", "align": "right"}],
    )


def claim(slide, num, text):
    tb(slide, ML, 0.4, 0.7, 0.36, [{"text": num, "font": MED, "size": 15, "color": RED}])
    tb(slide, 1.5, 0.34, 11, 0.48, [{"text": text, "font": MED, "size": 30, "color": INK}])
    rect(slide, ML, 1.05, 11.9, 0.012, HAIR)


def build():
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    prs.core_properties.title = "艾多美何以成为100兆韩元企业"
    prs.core_properties.subject = "蒙想讯息 · 2026年9月22日"
    total_note = "全稿 17 页"

    # 01 封面
    s = new(prs)
    bg(s, DARK)
    rect(s, 0, 0, 0.08, H, RED)
    tb(s, 0.85, 1.35, 10, 0.3, [{"text": "蒙想讯息    2026年9月22日", "font": SANS, "size": 14, "color": DIM}])
    tb(
        s, 0.82, 2.15, 12, 2.15,
        [
            {"text": "艾多美何以成为", "font": MED, "size": 48, "color": WHITE, "after": 0},
            {"text": "一百兆韩元企业", "font": MED, "size": 48, "color": WHITE},
        ],
    )
    rect(s, 0.88, 4.55, 1.05, 0.035, RED)
    tb(
        s, 0.88, 4.85, 11.2, 1.45,
        [{"text": "一百兆不是喊出来的。\n工序造出价格，组织代替店铺，选择被拿掉。\n中国法人要做的，是把这三件事在中国做成事实。", "font": SANS, "size": 18, "color": "D4D4D4", "line": 1.4}],
    )
    folio(s, 1, dark=True)
    notes(s, total_note + "。开场用这一句定标准：下面要证明的是结构，不是信念。")

    # 02 十倍
    s = new(prs)
    bg(s, WHITE)
    rect(s, 0, 0, 6.35, H, DARK)
    tb(s, 0.72, 1.55, 5, 0.28, [{"text": "行业第一", "font": SANS, "size": 14, "color": DIM}])
    tb(s, 0.68, 1.95, 5.3, 1.2, [{"text": "十兆", "font": MED, "size": 76, "color": WHITE}])
    tb(s, 0.72, 3.35, 5, 0.4, [{"text": "大约六十年", "font": MED, "size": 22, "color": WHITE}])
    tb(s, 0.72, 3.95, 5.1, 1.1, [{"text": "安利用了这么久，\n走到大约这个数字。", "font": SANS, "size": 16, "color": "BDBDBD", "line": 1.45}])
    tb(s, 6.95, 1.55, 5.6, 0.28, [{"text": "要去的地方", "font": SANS, "size": 14, "color": RED}])
    tb(s, 6.9, 1.95, 5.8, 1.2, [{"text": "一百兆", "font": MED, "size": 72, "color": RED}])
    tb(s, 6.95, 3.35, 5.6, 0.4, [{"text": "大约十五年", "font": MED, "size": 22, "color": INK}])
    tb(s, 6.95, 3.95, 5.6, 1.0, [{"text": "蒙想还在的年月。\n不是一百年之后。", "font": SANS, "size": 16, "color": BODY, "line": 1.45}])
    rect(s, 6.95, 5.35, 5.5, 0.012, HAIR)
    tb(s, 6.95, 5.6, 5.6, 1.15, [{"text": "追上第一，努力可以说通。\n十倍，必须有结构。", "font": MED, "size": 18, "color": INK, "line": 1.45}])
    folio(s, 2)
    notes(s, "先把落差放上桌。听众若不信，后面每一页都是写给这个不信的。")

    # 03 标准
    s = new(prs)
    bg(s, WHITE)
    claim(s, "01", "信念握不住一百兆")
    tb(s, ML, 1.4, 5.3, 0.45, [{"text": "信念", "font": MED, "size": 26, "color": INK}])
    tb(
        s, ML, 2.05, 5.2, 2.8,
        [{"text": "中心是自己。\n把“一定行”喊紧，念头会更硬。\n事情本身不会因此发生。", "font": SANS, "size": 18, "color": BODY, "line": 1.55}],
    )
    tb(s, ML, 5.3, 5.2, 1.2, [{"text": "口号只能把信念握紧。\n这一份稿要写给不信的人。", "font": SANS, "size": 16, "color": MUTED, "line": 1.45}])
    rect(s, 6.7, 1.28, 6.64, 5.55, DARK)
    tb(s, 7.15, 1.85, 5.6, 0.5, [{"text": "结构", "font": MED, "size": 26, "color": WHITE}])
    tb(
        s, 7.15, 2.55, 5.6, 3.2,
        [{"text": "中心是事情本身。\n信，或者不信，它都往那里去。\n下面要交出的，是这个“都”。", "font": SANS, "size": 18, "color": "E6E6E6", "line": 1.55}],
    )
    folio(s, 3)
    notes(s, "Belief 是把自己的想法握紧。要的是 Faith：对象如此，信不信都会发生。")

    # 04 市场
    s = new(prs)
    bg(s, WHITE)
    claim(s, "02", "一百兆不需要所有人")
    tb(
        s, ML, 1.3, 11.8, 0.7,
        [{"text": "亚马逊、沃尔玛、Coupang 把场子铺开，让人自己挑。教科书接着说：高低都做。\n这条正论一旦占了位置，绝对品质和绝对价格就站不稳。", "font": SANS, "size": 15, "color": BODY, "line": 1.4}],
    )
    rows = [
        ("丢掉", "最上约 1%", "嫌便宜。东西再好，也不用。", False),
        ("丢掉", "只认最便宜", "韩国大约三成。购买力更弱的地方，可以到七成。", False),
        ("留下", "绝对品质，绝对价格", "这一层拿下一半，就够走到千兆。不必占满。", True),
    ]
    y = 2.3
    for verb, who, why, hot in rows:
        rect(s, ML, y, 11.9, 0.01, HAIR)
        tb(s, ML, y + 0.22, 1.3, 0.4, [{"text": verb, "font": MED, "size": 16, "color": RED if hot else MUTED}])
        tb(s, 2.2, y + 0.16, 4.5, 0.5, [{"text": who, "font": MED, "size": 22, "color": RED if hot else INK}])
        tb(s, 7.0, y + 0.22, 5.5, 0.55, [{"text": why, "font": SANS, "size": 16, "color": BODY}])
        y += 1.35
    folio(s, 4)
    notes(s, "分母不是全部消费者。是同时要绝对品质和绝对价格的人，再取一半。")

    # 05 35%
    s = new(prs)
    bg(s, WHITE)
    claim(s, "03", "三十五个点，不能向顾客要")
    tb(s, ML, 1.35, 5.2, 1.25, [{"text": "35%", "font": MED, "size": 78, "color": RED}])
    tb(
        s, ML, 2.8, 5.3, 1.7,
        [{"text": "这是会员侧的津贴。\n现在的渠道大多只过一手。\n不能再解释成少了中间商。", "font": SANS, "size": 16, "color": BODY, "line": 1.45}],
    )
    specs = [
        ("好市多", "毛利为零。年费约八万韩元，买一千万和买一亿都是这一笔。生产者不另加负担。"),
        ("易买得", "进货比重常在六成到七成。"),
        ("电视购物", "播出费用大约四成二，生产者留下五成八。"),
    ]
    y = 1.35
    for i, (name, desc) in enumerate(specs):
        tb(s, 6.55, y, 6.1, 0.36, [{"text": name, "font": MED, "size": 18, "color": INK}])
        tb(s, 6.55, y + 0.4, 6.1, 0.75, [{"text": desc, "font": SANS, "size": 14, "color": BODY, "line": 1.3}])
        y += 1.25
        if i < 2:
            rect(s, 6.55, y - 0.12, 6.05, 0.01, HAIR)
    rect(s, ML, 5.35, 11.9, 0.012, HAIR)
    tb(
        s, ML, 5.55, 11.8, 1.15,
        [{"text": "向消费者加价，这一仗是输的。三十五个点要在工序里造出来。\n造出来之后，会员大约半价，事业者仍有津贴，公司仍有余量。", "font": SANS, "size": 16, "color": INK, "line": 1.45}],
    )
    folio(s, 5)
    notes(s, "好市多已经把毛利做成零。35% 若转嫁给顾客，就该劝会员去好市多。")

    # 06 牙刷
    s = new(prs)
    bg(s, WHITE)
    claim(s, "04", "一千八百韩元，不是利润")
    tb(s, ML, 1.35, 5.4, 1.15, [{"text": "1,200", "font": MED, "size": 72, "color": RED}])
    tb(s, ML, 2.6, 5.2, 0.7, [{"text": "韩元。市价 2,000 到 3,000。\n二十年，没有人跟上。", "font": SANS, "size": 16, "color": BODY, "line": 1.4}])
    layers = [
        ("制造", "型号一多，每样只做一点，成本被抬高。"),
        ("信任", "为了让人相信，广告和店铺把钱花出去。"),
        ("选择", "人要挑选，会买错，会退货。"),
    ]
    y = 1.4
    for name, desc in layers:
        tb(s, 6.7, y, 1.35, 0.4, [{"text": name, "font": MED, "size": 18, "color": RED}])
        tb(s, 8.15, y + 0.02, 4.4, 0.55, [{"text": desc, "font": SANS, "size": 16, "color": BODY}])
        y += 0.85
        rect(s, 6.7, y - 0.18, 5.85, 0.01, HAIR)
    tb(
        s, ML, 5.15, 11.8, 1.35,
        [{"text": "一个造型。人的嘴不变，模具就不改。腔数做大，原料买吨袋，货款付现金。\n若只是从现成的货里挑一个，Coupang 明天就能做。要做的是把这一件的生产从头重排。", "font": SANS, "size": 16, "color": INK, "line": 1.45}],
    )
    folio(s, 6)
    notes(s, "三层成本：多型号的制造、广告和店、消费者的选择失败。拿掉这三层，1,200 才成立。")

    # 07 HemoHIM
    s = new(prs)
    bg(s, WHITE)
    claim(s, "05", "七十七万，怎样落到七万六")
    rect(s, 0, 1.28, 6.35, 2.85, "F4F4F4")
    rect(s, 6.35, 1.28, 6.99, 2.85, DARK)
    tb(s, 0.72, 1.48, 5, 0.28, [{"text": "原先", "font": SANS, "size": 14, "color": MUTED}])
    tb(s, 0.68, 1.85, 5.3, 0.95, [{"text": "770,000", "font": MED, "size": 42, "color": INK}])
    tb(s, 0.72, 2.95, 5, 0.55, [{"text": "韩元。走干燥，走切断，\n在药材市场现买。", "font": SANS, "size": 14, "color": MUTED, "line": 1.35}])
    tb(s, 7.05, 1.48, 5.5, 0.28, [{"text": "现在", "font": SANS, "size": 14, "color": "E7B1B1"}])
    tb(s, 7.0, 1.82, 5.6, 1.05, [{"text": "76,500", "font": MED, "size": 48, "color": WHITE}])
    tb(s, 7.05, 2.95, 5.5, 0.55, [{"text": "韩元。价格停在十年前，\n品质走到十年后。", "font": SANS, "size": 14, "color": "E7B1B1", "line": 1.35}])
    facts = [
        ("0.5 吨到 18 吨", "三十六倍。操作的人仍是一个，人工按三十六分之一计。"),
        ("干燥与切断省掉", "大罐用洗净的鲜品直接取汁。药房那两道工序没有了。"),
        ("先锁住，再放量", "药材要一两年。合同栽培做在涨价之前。依据是月销十万盒。"),
    ]
    x = ML
    for i, (head, body) in enumerate(facts):
        if i:
            rect(s, x - 0.2, 4.5, 0.012, 1.7, HAIR)
        tb(s, x, 4.45, 3.6, 0.7, [{"text": head, "font": MED, "size": 16, "color": INK}])
        tb(s, x, 5.2, 3.55, 1.15, [{"text": body, "font": SANS, "size": 14, "color": BODY, "line": 1.4}])
        x += 4.05
    folio(s, 7)
    notes(s, "不是把价格砍下来。是把无用工序拿掉，再按一个月十万盒重算成本。")

    # 08 一贯工程
    s = new(prs)
    bg(s, WHITE)
    claim(s, "06", "重做的是整条线")
    cols = [
        ("01", "走到原料", "人们以为搅拌的人在制造。真正的条件在原料厂：温度、压力、配比。设备厂手里有运转法。按一年的量去要图纸。"),
        ("02", "涨价之前锁住", "工业品可以连夜加产。药材不行。工序要先放到不能再省的地方。绝对价格，是谁来做都无法更便宜。"),
        ("03", "会做，才交出运转", "设备我们出，原料我们买，对方只运转。一起，两亿的公司可以做到一百亿。不一起，自己做。"),
    ]
    x = ML
    for num, head, body in cols:
        tb(s, x, 1.4, 3.6, 0.3, [{"text": num, "font": MED, "size": 13, "color": RED}])
        tb(s, x, 1.8, 3.65, 0.85, [{"text": head, "font": MED, "size": 24, "color": INK}])
        tb(s, x, 2.8, 3.6, 2.5, [{"text": body, "font": SANS, "size": 15, "color": BODY, "line": 1.45}])
        x += 4.05
    rect(s, ML, 5.7, 11.9, 0.012, HAIR)
    tb(s, ML, 5.95, 11.8, 0.75, [{"text": "洗发水也是这条路：并成一个，换料和清洗的损耗去掉，成本至少降一半。\n一百兆不附带“对方肯配合”。跟，就一起。不跟，优化也不停。", "font": SANS, "size": 15, "color": INK, "line": 1.4}])
    folio(s, 8)
    notes(s, "外包的意思变了：不是自己不会，是运转属于较低的一层，所以交出去。系统留在自己手里。")

    # 09 大创
    s = new(prs)
    bg(s, WHITE)
    claim(s, "07", "生产侧，对方已经走到了")
    tb(s, ML, 1.45, 5.2, 0.3, [{"text": "大创的销售", "font": SANS, "size": 14, "color": MUTED}])
    tb(s, ML - 0.02, 1.85, 5.4, 1.1, [{"text": "4.5 兆", "font": MED, "size": 60, "color": INK}])
    tb(s, ML, 3.15, 5.3, 1.3, [{"text": "增长维持在三成以上。\n策展、放量、进产线，它都做。\n便宜，不等于粗糙。", "font": SANS, "size": 16, "color": BODY, "line": 1.45}])
    rect(s, 6.55, 1.55, 0.012, 2.9, HAIR)
    tb(s, 7.0, 1.45, 5.4, 0.3, [{"text": "艾多美眼下", "font": SANS, "size": 14, "color": RED}])
    tb(s, 6.96, 1.85, 5.5, 1.1, [{"text": "约 3%", "font": MED, "size": 60, "color": RED}])
    tb(s, 7.0, 3.15, 5.4, 1.3, [{"text": "这是增长，不是销售。\n品质更好、设计更好，\n解释不了这个差距。", "font": SANS, "size": 16, "color": BODY, "line": 1.45}])
    rect(s, ML, 4.85, 11.9, 0.012, HAIR)
    tb(
        s, ML, 5.15, 11.8, 1.4,
        [{"text": "生产上能挖到的，对方大多已经挖到。差别只能到销售这一侧去找。\n若邻里觉得街上有差不多的东西，价格只有一半的一半，后面的教育再密，话也传不出去。", "font": SANS, "size": 16, "color": INK, "line": 1.45}],
    )
    folio(s, 9)
    notes(s, "大创比一般卖场更要认真。经营者从制造里出来。品牌更高，是不够的答案。")

    # 10 组织
    s = new(prs)
    bg(s, WHITE)
    claim(s, "08", "没有店，那三十五个点才回得来")
    tb(s, ML, 1.35, 5.4, 0.4, [{"text": "大创", "font": MED, "size": 22, "color": INK}])
    tb(s, ML, 1.9, 5.4, 1.35, [{"text": "必须先有店，人再来。\n店的成本，和人找上门的成本，\n都是实的。没有店，就没有消费者。", "font": SANS, "size": 16, "color": BODY, "line": 1.45}])
    rect(s, 6.7, 1.28, 6.64, 2.35, DARK)
    tb(s, 7.05, 1.48, 5.7, 0.4, [{"text": "艾多美", "font": MED, "size": 22, "color": WHITE}])
    tb(s, 7.05, 2.0, 5.8, 1.3, [{"text": "没有店。人自己用，也介绍。\n半价的体验代替广告和店铺。\n省下来的，就是那 35%。", "font": SANS, "size": 16, "color": "E8E8E8", "line": 1.4}])
    steps = ["绝对价格", "会员", "产量", "成本再降"]
    for i, name in enumerate(steps):
        x = ML + i * 3.0
        tb(s, x, 4.0, 2.85, 0.45, [{"text": name, "font": MED, "size": 18, "color": RED if i == 0 else INK, "align": "center"}])
    rect(s, ML, 4.6, 11.9, 0.012, HAIR)
    tb(
        s, ML, 4.85, 11.8, 1.6,
        [{"text": "价格带来会员，会员堆出产量，产量把成本再压低。轮子没有刹车。\n摩洛哥会员十万以上，孟加拉约二十万。收入差这么远，卖的是同一套产品。\n这种东西不是广告推过境的。组织先动，店可以后有。", "font": SANS, "size": 16, "color": BODY, "line": 1.45}],
    )
    folio(s, 10)
    notes(s, "非洲、孟加拉、秘鲁、巴拿马，都是队伍先走。店不是起点。")

    # 11 策展
    s = new(prs)
    bg(s, WHITE)
    claim(s, "09", "策展不停在货架上")
    qs = [
        ("留下什么", "一个品类，一个。品质不拿来换价格。顾客不必在几百个相近的东西里比较。"),
        ("如何守住", "从原物管到流通。工厂不必座座自有。原料、设备基准和数据要在自己手里。循环的第一步，是对方看得见的愿景，不是先把货堆出来。"),
        ("对准谁", "同一种产品，不再对所有人说同一句话。身体不同，建议就不同。"),
    ]
    y = 1.35
    for head, body in qs:
        tb(s, ML, y, 2.7, 0.7, [{"text": head, "font": MED, "size": 20, "color": INK}])
        tb(s, 3.6, y + 0.02, 8.9, 0.85, [{"text": body, "font": SANS, "size": 16, "color": BODY, "line": 1.35}])
        y += 1.15
        rect(s, ML, y - 0.18, 11.9, 0.01, HAIR)
    tb(
        s, ML, 5.15, 11.8, 1.2,
        [{"text": "所以五百个，不能按五百个算。\n亚马逊一把牙刷就有几百款。我们一把，顶的是被挑剩的那一个。", "font": MED, "size": 18, "color": RED, "line": 1.4}],
    )
    folio(s, 11)
    notes(s, "开放市场用长尾。艾多美把量集中到一个做得好的合作方。愿景先被看见，大量生产的循环才进得去。")

    # 12 三A
    s = new(prs)
    bg(s, WHITE)
    claim(s, "10", "三个工具，是一个结构")
    engines = [
        ("APP", "个人平台", "自己的店、自己的健康、自己的教育，收在一处。过去租酒店讲一小时，收成五分钟。平台本身就是人工智能。"),
        ("A-Care", "艾护理", "不再说“做好了请用”。血、饮食、活动放进来，一年、五年、十年后的身体可以估。越用越准，人才留下来。"),
        ("AZA", "阿扎", "鸡蛋、油、加油接进来，负责人每天走进来。十万个品目不赚一百兆。一百兆，由大约五百个主力品目赚。"),
    ]
    x = ML
    for i, (en, cn, body) in enumerate(engines):
        if i:
            rect(s, x - 0.22, 1.4, 0.012, 3.35, HAIR)
        tb(s, x, 1.35, 3.55, 0.28, [{"text": en, "font": MED, "size": 13, "color": RED}])
        tb(s, x, 1.75, 3.6, 0.5, [{"text": cn, "font": MED, "size": 24, "color": INK}])
        tb(s, x, 2.45, 3.55, 2.2, [{"text": body, "font": SANS, "size": 15, "color": BODY, "line": 1.45}])
        x += 4.05
    rect(s, ML, 5.15, 11.9, 0.012, HAIR)
    tb(
        s, ML, 5.4, 11.8, 1.15,
        [{"text": "事业者因此不再只是推销的人。工具一变，行为就变，身份就变成顾问，变成这家平台的主人。\n消费、推荐和组织，变成留得下的资产。人留下，产量才留下。", "font": SANS, "size": 16, "color": INK, "line": 1.4}],
    )
    folio(s, 12)
    notes(s, "艾护理负责关系，阿扎负责每天来，个人平台负责把事业做成方法。加油走联名卡积分，不卖加油券。")

    # 13 四句
    s = new(prs)
    bg(s, WHITE)
    claim(s, "11", "四件事情要同时为真")
    lines = [
        ("01", "这条链不容易被照抄", "价格表好抄。育苗、栽培、图纸、专用线和数据，要有人先把愿景和钱放进去。"),
        ("02", "不附带对方肯配合", "跟，就一起把小公司做大。不跟，优化也不停。目标没有这个条件。"),
        ("03", "一半，就是千兆的底", "不要全部市场。绝对品质、绝对价格这一层的一半，已经够。"),
        ("04", "人留下，产量才留下", "健康越用越准，日常开支接在一处，内容和组织长在自己的平台上。"),
    ]
    y = 1.35
    for num, head, body in lines:
        tb(s, ML, y, 0.7, 0.4, [{"text": num, "font": MED, "size": 16, "color": RED}])
        tb(s, 1.6, y - 0.02, 10.8, 0.4, [{"text": head, "font": MED, "size": 22, "color": INK}])
        tb(s, 1.6, y + 0.46, 10.8, 0.4, [{"text": body, "font": SANS, "size": 15, "color": BODY}])
        y += 1.3
    folio(s, 13)
    notes(s, "可以分开讲。不能分开成立。四句同时转，才是一百兆。")

    # 14 中国的三种诱惑
    s = new(prs)
    bg(s, WHITE)
    claim(s, "12", "中国把三种诱惑同时放大")
    traps = [
        ("01", "自己挑", "天猫、京东把货铺开，让人在里面比较。选择一多，绝对品质就松。"),
        ("02", "只要便宜", "低价平台把“最贱”做成默认。\n只认便宜的那一层，在中国更响。\n跟过去，就会另做一条贱的线。"),
        ("03", "先把店开起来", "开市客、大创、名创优品，\n都用店把人接住。\n中国法人若先开店，三十五个点就没了。"),
    ]
    x = ML
    for i, (num, head, body) in enumerate(traps):
        if i:
            rect(s, x - 0.22, 1.4, 0.012, 3.15, HAIR)
        tb(s, x, 1.4, 3.5, 0.3, [{"text": num, "font": MED, "size": 14, "color": RED}])
        tb(s, x, 1.85, 3.55, 0.85, [{"text": head, "font": MED, "size": 26, "color": INK}])
        tb(s, x, 2.9, 3.55, 1.7, [{"text": body, "font": SANS, "size": 15, "color": BODY, "line": 1.45}])
        x += 4.05
    rect(s, ML, 5.15, 11.9, 0.012, HAIR)
    tb(
        s, ML, 5.45, 11.8, 1.2,
        [{"text": "三种都在叫中国法人改模式。\n改了，就不再是同一套结构。", "font": MED, "size": 22, "color": RED, "line": 1.35}],
    )
    folio(s, 14)
    notes(s, "中国市场把细分、低价和门店三句同时喊响。中国法人的第一件事，是这三句都不跟。")

    # 15 中国法人四件
    s = new(prs)
    bg(s, WHITE)
    claim(s, "13", "中国法人就做这四件")
    duties = [
        ("01", "不另做一条中国线", "一个品类仍是一个。不为最上的人做贵的，不为只认便宜的人做贱的。"),
        ("02", "不在中国价格上加一层", "三十五个点不向中国消费者要。加了，就该让人去开市客。"),
        ("03", "用中国的量，把工序锁住", "把看得见的需求交给一条线。愿景让对方看见，比先堆很多款有用。"),
        ("04", "先有队伍，再谈店和广告", "店和投放吃掉的，就是那三十五个点。组织先动，店可以后有。"),
    ]
    y = 1.32
    for num, head, body in duties:
        tb(s, ML, y, 0.7, 0.4, [{"text": num, "font": MED, "size": 16, "color": RED}])
        tb(s, 1.6, y - 0.02, 10.8, 0.42, [{"text": head, "font": MED, "size": 22, "color": INK}])
        tb(s, 1.6, y + 0.46, 10.8, 0.4, [{"text": body, "font": SANS, "size": 15, "color": BODY}])
        y += 1.28
    folio(s, 15)
    notes(s, "四件都是同一套结构在中国的动作。加款、加价、加店、加广告，都是把方向做松。")

    # 16 三台引擎在中国
    s = new(prs)
    bg(s, WHITE)
    claim(s, "14", "三台引擎，在中国转起来")
    local = [
        ("个人平台", "用中国会员听得懂的话，\n把一小时收成五分钟。\n事业者是这家平台的主人，\n不是发链接的人。"),
        ("艾护理", "从身体和饮食进入，\n不再“好东西请用”。\n人越用越留下，建议才因人而异。\n不必等一座医院盖好才开始。"),
        ("阿扎", "油、蛋、充电，每天的开支接进来。\n电要在规则里面找能累积的办法。\n十万个日常品目让人每天来。\n一百兆仍由大约五百个主力品目赚。"),
    ]
    x = ML
    for i, (head, body) in enumerate(local):
        if i:
            rect(s, x - 0.22, 1.4, 0.012, 3.15, HAIR)
        tb(s, x, 1.45, 3.55, 0.6, [{"text": head, "font": MED, "size": 24, "color": INK}])
        tb(s, x, 2.3, 3.55, 2.2, [{"text": body, "font": SANS, "size": 15, "color": BODY, "line": 1.45}])
        x += 4.05
    rect(s, ML, 5.15, 11.9, 0.012, HAIR)
    tb(
        s, ML, 5.45, 11.8, 1.2,
        [{"text": "中国法人不另造一个中国模式。\n同一套结构，在中国转起来。", "font": MED, "size": 22, "color": RED, "line": 1.35}],
    )
    folio(s, 16)
    notes(s, "工具不换一套。换的是中国会员听得懂的话，和规则里面能做成的接法。医院不是开工的前提。")

    # 17 收束
    s = new(prs)
    bg(s, DARK)
    rect(s, 0, 0, 0.08, H, RED)
    lines = [
        ("生产被重做", "三十五个点有了出处"),
        ("消费者被组织", "广告和店铺被省掉"),
        ("选择被拿掉", "一个品类，只留一个"),
        ("人留在结构里", "产量才继续把成本压低"),
    ]
    y = 0.72
    for a, b in lines:
        tb(s, 0.85, y, 5.0, 0.5, [{"text": a, "font": MED, "size": 26, "color": WHITE}])
        tb(s, 6.15, y + 0.08, 6.3, 0.4, [{"text": b, "font": SANS, "size": 16, "color": "B5B5B5"}])
        y += 0.78
    tb(s, 0.82, 4.05, 11, 0.85, [{"text": "不能不去", "font": MED, "size": 48, "color": RED}])
    tb(s, 0.88, 5.15, 11, 0.7, [{"text": "中国法人要做的，是把这四句在中国做成事实。", "font": SANS, "size": 18, "color": "D4D4D4"}])
    tb(s, 0.88, 6.35, 9, 0.3, [{"text": "蒙想讯息    2026年9月22日", "font": SANS, "size": 14, "color": DIM}])
    folio(s, 17, dark=True)
    notes(s, "收束不再加新论点。四句对上封面。最后一句落到中国法人自己身上，然后停住。")

    return prs


if __name__ == "__main__":
    path = "/workspace/艾多美何以成为100兆韩元企业.pptx"
    build().save(path)
    print(path)
