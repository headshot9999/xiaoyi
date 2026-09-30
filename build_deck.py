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
    total_note = "全稿 16 页"

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
        [{"text": "一个功能，一个产品。品质加倍，或价格减半。\n成本从起点就按超大规模算。\n写给中国法人：哪些工序拿掉，三 A 怎么接。", "font": SANS, "size": 18, "color": "D4D4D4", "line": 1.4}],
    )
    folio(s, 1, dark=True)
    notes(s, total_note + "。只讲中国法人要做的事：标准、产线、合作方、企划、三 A。")

    # 02 三条标准
    s = new(prs)
    bg(s, WHITE)
    claim(s, "01", "回到重定标准的时候")
    standards = [
        ("01", "一能一品", "散开的品类收成一个。\n一个功能，一个产品。\n这一个，要成为品类里的\n全球冠军。\n中国的量用在这一个上。"),
        ("02", "加倍，或减半", "品质加倍，或价格减半。\n两件可以一起发生。\n品质是世界第一。\n价格是压倒性的第一。"),
        ("03", "成本先按超大算", "从起点就按超大规模重算。\n一个月十万套是前提，\n不是卖起来再去谈。\n谈运费，谁都会。"),
    ]
    x = ML
    for i, (num, head, body) in enumerate(standards):
        if i:
            rect(s, x - 0.22, 1.32, 0.012, 3.35, HAIR)
        tb(s, x, 1.32, 3.5, 0.26, [{"text": num, "font": MED, "size": 14, "color": RED}])
        tb(s, x, 1.64, 3.55, 0.7, [{"text": head, "font": MED, "size": 22, "color": INK}])
        tb(s, x, 2.45, 3.55, 2.2, [{"text": body, "font": SANS, "size": 15, "color": BODY, "line": 1.35}])
        x += 4.05
    rect(s, ML, 5.2, 11.9, 0.012, HAIR)
    tb(
        s, ML, 5.42, 11.8, 1.15,
        [{"text": "总目标：十年前的价格，十年后的品质。\n按规格跟着市场走，标准就在对方手里。", "font": MED, "size": 20, "color": RED, "line": 1.3}],
    )
    folio(s, 2)
    notes(s, "7月28日工作坊：面对按规格生产的竞争者，回到重新制定标准。三条是一能一品、品质加倍或价格减半、从超大规模重构成本。9月13日把品质定为世界第一、价格竞争力定为压倒性第一，一能一品是一个产品、一个全球冠军。")

    # 03 牙刷
    s = new(prs)
    bg(s, WHITE)
    claim(s, "02", "一千八百韩元，不是利润")
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
    folio(s, 3)
    notes(s, "三层成本：多型号的制造、广告和店、消费者的选择失败。拿掉这三层，1,200 才成立。")

    # 04 HemoHIM
    s = new(prs)
    bg(s, WHITE)
    claim(s, "03", "七十七万，怎样落到七万六")
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
    folio(s, 4)
    notes(s, "不是把价格砍下来。是把无用工序拿掉，再按一个月十万盒重算成本。")

    # 05 产线上的四件
    s = new(prs)
    bg(s, WHITE)
    claim(s, "04", "中国产线上就做这四件")
    duties = [
        ("01", "一个模具做大", "牙刷造型不改。腔数做大，机器换大，原料买吨袋，付现金。再换一个造型，腔数就做不大。"),
        ("02", "一种配方，罐子不洗", "洗发水四五款并成一款。换料和清洗都在多款里。再加一个口味，清洗就回来。"),
        ("03", "罐子做大，人还是一个", "萃取罐从 0.5 吨做到 18 吨，三十六倍，操作的人仍是一个。要推的是这只罐子，不是多找几家小厂。"),
        ("04", "鲜品进罐", "川芎、当归、白芍洗净直接煮。干燥和切断是按克卖药才有的。中国法人在产地守的是这一步。"),
    ]
    y = 1.28
    for num, head, body in duties:
        tb(s, ML, y, 0.7, 0.36, [{"text": num, "font": MED, "size": 16, "color": RED}])
        tb(s, 1.55, y - 0.02, 10.9, 0.4, [{"text": head, "font": MED, "size": 22, "color": INK}])
        tb(s, 1.55, y + 0.42, 10.9, 0.7, [{"text": body, "font": SANS, "size": 15, "color": BODY}])
        y += 1.32
    rect(s, ML, 6.3, 11.9, 0.012, HAIR)
    tb(
        s, ML, 6.42, 11.8, 0.5,
        [{"text": "谈运费，谁都会。这四件，只有一贯工程才会做。", "font": MED, "size": 18, "color": RED}],
    )
    folio(s, 5)
    notes(s, "四件都落在工序上：模具、配方、罐容、鲜品。加款、谈运费、再加一个口味，都不在这四件里。")

    # 06 烟台自查单
    s = new(prs)
    bg(s, WHITE)
    claim(s, "05", "烟台先查这四条线")
    tb(s, ML, 1.25, 3.0, 0.3, [{"text": "线", "font": SANS, "size": 13, "color": MUTED}])
    tb(s, 3.7, 1.25, 5.0, 0.3, [{"text": "查什么", "font": SANS, "size": 13, "color": MUTED}])
    tb(s, 9.55, 1.25, 3.2, 0.3, [{"text": "交出什么", "font": SANS, "size": 13, "color": MUTED}])
    rows = [
        ("牙刷", "烟台工厂造", "型号几款；模具几腔；\n原料是袋装还是吨袋；\n付现金还是赊账。", "缺哪几项，\n按总部标准列出。"),
        ("厨具", "烟台工厂造", "同一个功能几款；\n换款一次，停线多久。", "并款清单。"),
        ("保健食品", "烟台工厂造", "萃取罐几吨，几个人操作；\n干燥、切断两步还在不在。", "鲜品直接取汁的\n改造单。"),
        ("物流", "烟台国际物流中心，2025 年启用", "原料到下一道工序，中间转几趟；\n下一道工序能不能放到\n原料厂的墙外。", "少一趟的清单。"),
    ]
    y = 1.62
    for name, sub, check, out in rows:
        rect(s, ML, y - 0.08, 11.9, 0.01, HAIR)
        tb(s, ML, y, 2.9, 0.4, [{"text": name, "font": MED, "size": 20, "color": INK}])
        tb(s, ML, y + 0.42, 2.9, 0.5, [{"text": sub, "font": SANS, "size": 12, "color": MUTED}])
        tb(s, 3.7, y, 5.6, 0.95, [{"text": check, "font": SANS, "size": 15, "color": BODY, "line": 1.3}])
        tb(s, 9.55, y, 3.2, 0.95, [{"text": out, "font": SANS, "size": 15, "color": BODY, "line": 1.3}])
        y += 1.07
    rect(s, ML, y - 0.08, 11.9, 0.01, HAIR)
    tb(s, ML, 6.0, 11.8, 0.45, [{"text": "先各自盘点，再谈改。", "font": MED, "size": 20, "color": RED}])
    tb(s, ML, 6.6, 11.8, 0.3, [{"text": "烟台工厂与物流中心的情况引自《艾多美 100 兆亿养成手册》，需核实。", "font": SANS, "size": 11, "color": MUTED}])
    folio(s, 6)
    notes(s, "这一页不写结论，只列要查的。四条线按四件事对应：牙刷是模具和吨袋，厨具是并款，保健食品是罐容和鲜品，物流是转运趟数。烟台工厂造厨具、牙刷、保健食品，以及烟台国际物流中心2025年启用，均引自手册，未核实。")

    # 07 带去合作方的三样
    s = new(prs)
    bg(s, WHITE)
    claim(s, "06", "带去合作方的，是这三样")
    local = [
        ("一个月十万套", "价格按这个量重算，\n才落到原来的十分之一。\n交出去的是能排产的这个数。"),
        ("设备图纸", "设备公司出。没有厂房，就在旁边盖。\n原料买断，对方只运转。\n拿去的是图纸，不是询价单。"),
        ("育苗就开始的合同", "川芎、当归、白芍从育苗起要一两年。\n合同在涨价之前锁上。\n农活不会，就公司带。"),
    ]
    x = ML
    for i, (head, body) in enumerate(local):
        if i:
            rect(s, x - 0.22, 1.4, 0.012, 3.15, HAIR)
        tb(s, x, 1.5, 3.55, 1.15, [{"text": head, "font": MED, "size": 24, "color": INK}])
        tb(s, x, 2.8, 3.55, 1.8, [{"text": body, "font": SANS, "size": 15, "color": BODY, "line": 1.45}])
        x += 4.05
    rect(s, ML, 5.15, 11.9, 0.012, HAIR)
    tb(
        s, ML, 5.45, 11.8, 1.2,
        [{"text": "大量生产不是第一步。\n这个数、这张图纸、这份合同，才是。", "font": MED, "size": 22, "color": RED, "line": 1.35}],
    )
    folio(s, 7)
    notes(s, "合作方先要看见能排产的量、设备由谁出、原料从育苗锁到哪一年。询价单给不出这些。")

    # 08 不能接力
    s = new(prs)
    bg(s, WHITE)
    claim(s, "07", "从企划一次通到终点")
    relay = [
        ("01", "接力做不成一炉", "商品做完才做内容，内容做完才培训。任何一环停下，后面都在等。"),
        ("02", "中间不能熄火", "钢铁厂从原料到成品不能中断。干燥、切断、运去别处再运回，都是中间熄火。"),
        ("03", "一次写完", "这个产品、这条工序、中文骨架、对这个人的那句话、每天打开的那一件，写在同一次企划里。"),
    ]
    y = 1.35
    for num, head, body in relay:
        tb(s, ML, y, 0.7, 0.36, [{"text": num, "font": MED, "size": 16, "color": RED}])
        tb(s, 1.55, y - 0.02, 10.9, 0.4, [{"text": head, "font": MED, "size": 22, "color": INK}])
        tb(s, 1.55, y + 0.46, 10.9, 0.55, [{"text": body, "font": SANS, "size": 16, "color": BODY}])
        y += 1.22
    rect(s, ML, 5.2, 11.9, 0.012, HAIR)
    tb(
        s, ML, 5.45, 11.8, 1.1,
        [{"text": "中国法人先出货、再补内容、再办培训，\n链就在企划上断了。", "font": MED, "size": 22, "color": RED, "line": 1.3}],
    )
    folio(s, 8)
    notes(s, "7月28日：批评线性接力，要求从一体化企划一次贯通到终点。中国法人先出货、再补内容、再办培训，链在企划上就断了。")

    # 09 个人平台
    s = new(prs)
    bg(s, WHITE)
    claim(s, "08", "个人平台，先改变事业者")
    app = [
        ("01", "八个阶段在里面", "产品懂得多少，\n八个阶段走到哪一步，\n下一步学什么，都在里面。\n中国法人做成中文的路径。\n人不用等一场酒店课。"),
        ("02", "一百到二百支骨架", "到明年年底先打好骨架。\n事业者换成自己的脸、\n自己的讲法。\n一小时收成五分钟。\n说明在这个人的平台上完成。"),
        ("03", "PPO，平台的主人", "PPO 是这家平台的主人，\n不是发链接的人。\n自己的店、健康、教育，\n收在一处。\n培育 PPO，是九月十三日\n三份发表之一。"),
    ]
    x = ML
    for i, (num, head, body) in enumerate(app):
        if i:
            rect(s, x - 0.22, 1.32, 0.012, 3.35, HAIR)
        tb(s, x, 1.32, 3.5, 0.26, [{"text": num, "font": MED, "size": 14, "color": RED}])
        tb(s, x, 1.64, 3.55, 0.7, [{"text": head, "font": MED, "size": 22, "color": INK}])
        tb(s, x, 2.45, 3.55, 2.15, [{"text": body, "font": SANS, "size": 15, "color": BODY, "line": 1.35}])
        x += 4.05
    rect(s, ML, 5.2, 11.9, 0.012, HAIR)
    tb(
        s, ML, 5.42, 11.8, 1.15,
        [{"text": "教育的大笔费用在这里消失。\n省下来的，还是那三十五个点。", "font": MED, "size": 22, "color": RED, "line": 1.3}],
    )
    folio(s, 9)
    notes(s, "个人平台先武装事业者：阶段、中文骨架、PPO。培育 PPO 是9月13日奖项的三份发表之一。酒店课和培训部是旧成本。")

    # 10 艾护理
    s = new(prs)
    bg(s, WHITE)
    claim(s, "09", "艾护理对着这个人说")
    tb(s, ML, 1.28, 5.5, 0.4, [{"text": "放进来", "font": MED, "size": 20, "color": RED}])
    tb(s, 6.85, 1.28, 5.6, 0.4, [{"text": "说出来", "font": MED, "size": 20, "color": RED}])
    ins = [
        ("每餐一张照片", "碳水、蛋白、脂肪，白米还是玄米，蔬菜和纤维有多少。"),
        ("身高、体重、步数", "肥胖还是肌肉，活动够不够。步数从手机接进来。"),
        ("用药和注射", "代谢顺不顺，有连续记录才说得清。"),
        ("血和尿可以后补", "上面这些先每天跑起来。不必等一座医院。"),
    ]
    outs = [
        ("该减，还是该补", "减碳水，补蛋白；吃多了；或只吃精制，碳水其实不够。"),
        ("一年、五年、十年", "照这样下去，身体可以估。家族里带下来的病要能看见。"),
        ("缺哪个功能", "一能一品：先说缺哪个功能，那一个产品跟着定。"),
        ("先进入艾护理", "介绍别人，先把对方的记录建起来，再谈产品。"),
    ]
    y = 1.72
    for head, body in ins:
        tb(s, ML, y, 5.5, 0.3, [{"text": head, "font": MED, "size": 16, "color": INK}])
        tb(s, ML, y + 0.3, 5.5, 0.42, [{"text": body, "font": SANS, "size": 14, "color": BODY}])
        y += 0.76
    y = 1.72
    for head, body in outs:
        tb(s, 6.85, y, 5.6, 0.3, [{"text": head, "font": MED, "size": 16, "color": INK}])
        tb(s, 6.85, y + 0.3, 5.6, 0.42, [{"text": body, "font": SANS, "size": 14, "color": BODY}])
        y += 0.76
    rect(s, 6.45, 1.28, 0.012, 3.5, HAIR)
    rect(s, ML, 5.05, 11.9, 0.012, HAIR)
    tb(
        s, ML, 5.22, 11.8, 1.15,
        [{"text": "终点是个人健康数据银行。本人同意，隐私守住。\n中国要先定两件：记录存在哪里，哪句话能说。", "font": MED, "size": 18, "color": RED, "line": 1.3}],
    )
    folio(s, 10)
    notes(s, "艾护理的输入是每餐、体型、步数、用药。输出先是这个人缺哪个功能，产品因一能一品而跟着定，人工智能不是在几个产品里挑。血和尿可以后补，记录不能等医院。记录存哪里、哪句话能说，中国要先定，定之前不对会员说。")

    # 11 阿扎
    s = new(prs)
    bg(s, WHITE)
    claim(s, "10", "阿扎让人每天走进来")
    aza = [
        ("01", "十万个，不赚一百兆", "鸡蛋、油、充电接进来。\n一百兆由大约五百个\n主力品目赚。\n阿扎做的是热狗那一件事。\n先接每天都发生的开支。"),
        ("02", "鲭鱼回答不了", "成本率到七成五的日常品，\n单靠它到不了一百兆。\n它负责人每天来。\n主力品目负责赚。"),
        ("03", "电，先定哪张卡", "不卖加油券，\n走联名卡积分。\n中国电动车多。\n用哪张卡、哪张充电网，\n先定。能累积，再对会员说。"),
    ]
    x = ML
    for i, (num, head, body) in enumerate(aza):
        if i:
            rect(s, x - 0.22, 1.32, 0.012, 3.35, HAIR)
        tb(s, x, 1.32, 3.5, 0.26, [{"text": num, "font": MED, "size": 14, "color": RED}])
        tb(s, x, 1.64, 3.55, 0.7, [{"text": head, "font": MED, "size": 22, "color": INK}])
        tb(s, x, 2.45, 3.55, 2.15, [{"text": body, "font": SANS, "size": 15, "color": BODY, "line": 1.35}])
        x += 4.05
    rect(s, ML, 5.2, 11.9, 0.012, HAIR)
    tb(
        s, ML, 5.42, 11.8, 1.15,
        [{"text": "每天的开支接进同一个平台。\n人每天来，五百个主力的产量才留得住。", "font": MED, "size": 20, "color": RED, "line": 1.3}],
    )
    folio(s, 11)
    notes(s, "阿扎是市集的地面，不是行政窗口。鲭鱼这类高成本率的日常品负责人来。一百兆仍由大约五百个主力品目赚。充电接不接、走哪张卡哪张网，中国先定，不卖加油券。")

    # 12 一条线
    s = new(prs)
    bg(s, WHITE)
    claim(s, "11", "三台是一条线，不是三个栏")
    chain = [
        ("01", "艾护理建记录", "新人先进来。\n每餐、体型、步数、\n用药。\n记录先每天跑起来。"),
        ("02", "说缺什么", "先说这个人\n缺哪个功能。\n一个功能，\n只有一个产品。"),
        ("03", "那一个", "进的就是品类里\n那一个冠军。\n不必再比较，\n话对着这个人说。"),
        ("04", "阿扎每天", "每天的开支接进来。\n人每天打开。\n主力品目在这里\n被带出来。"),
        ("05", "PPO 与排产", "用的人记在 PPO 名下。\n合起来，是下个月\n那只罐子和栽培合同\n看得见的数。"),
    ]
    x = ML
    for i, (num, head, body) in enumerate(chain):
        if i:
            rect(s, x - 0.13, 1.32, 0.012, 3.35, HAIR)
        tb(s, x, 1.32, 2.2, 0.26, [{"text": num, "font": MED, "size": 14, "color": RED}])
        tb(s, x, 1.64, 2.3, 0.7, [{"text": head, "font": MED, "size": 20, "color": INK}])
        tb(s, x, 2.45, 2.3, 2.2, [{"text": body, "font": SANS, "size": 14, "color": BODY, "line": 1.35}])
        x += 2.42
    rect(s, ML, 5.05, 11.9, 0.012, HAIR)
    tb(
        s, ML, 5.25, 11.8, 1.3,
        [{"text": "三台是一次企划的一条线。\n要数三样：建了记录的人，连续用的天数，每天打开阿扎的人。", "font": MED, "size": 20, "color": RED, "line": 1.3}],
    )
    folio(s, 12)
    notes(s, "三 A 不是三个并列的应用。新人进艾护理建记录，记录说缺哪个功能，一能一品定那一个产品，阿扎每天接开支带出主力品目，用的人记在 PPO 名下，合起来变成能排产的数。要数的三样由中国法人自己给出，这里不填数字。")

    # 13 商城
    s = new(prs)
    bg(s, WHITE)
    claim(s, "12", "首页两样，语音只确认")
    mall = [
        ("01", "千人千面", "打开不是目录。今天只摆两样：缺的那个功能对应的产品，旁边一句为什么；再加上每天用的那一件。人不同，两样不同。功效句只用审过的。没定哪句话能说，首页不新写。"),
        ("02", "推荐只做三件", "缺什么：记录决定功能，功能决定那一个，不另推替代款。快用完了：按上次购买和用量提醒续上。该带给谁：先发进入艾护理的入口，不发产品链接。"),
        ("03", "语音加购与支付", "说“把今天这两样加入购物车”，再说“确认支付”。语音不搜索、不比较、不改款。支付走现有收银台，语音里不报卡号、不存卡号。收银台没接上之前，语音只加购。"),
        ("04", "下完单，三处一起变", "记在 PPO 名下。这个人的连续使用更新。同类订单合成下个月那只罐子要排的数。事业者工作台只看三行：谁还没建记录，谁的瓶子快空了，谁停了几天。"),
    ]
    y = 1.22
    for num, head, body in mall:
        tb(s, ML, y, 0.55, 0.32, [{"text": num, "font": MED, "size": 14, "color": RED}])
        tb(s, 1.35, y - 0.02, 2.5, 0.36, [{"text": head, "font": MED, "size": 18, "color": INK}])
        tb(s, 3.9, y, 8.7, 0.85, [{"text": body, "font": SANS, "size": 14, "color": BODY, "line": 1.25}])
        y += 1.05
        rect(s, ML, y - 0.12, 11.9, 0.01, HAIR)
    tb(
        s, ML, 5.55, 11.8, 1.0,
        [{"text": "平台已经会在几百款里推荐。\n中国法人的商城不做那一件。", "font": MED, "size": 20, "color": RED, "line": 1.25}],
    )
    folio(s, 13)
    notes(s, "千人千面是两样和那句话因人而异，不是目录变长。推荐只做缺什么、快用完、该带给谁。语音只确认已经摆上的两样，支付走现有收银台，不在语音里收集卡号。下单同时记到 PPO、连续使用和排产。")

    # 14 三步
    s = new(prs)
    bg(s, WHITE)
    claim(s, "13", "三步：做什么，交什么，谁牵头")
    steps3 = [
        ("01", "现在起", "烟台四条线自查。\n定两件事：记录存哪里、\n哪句话能说。\n定阿扎接哪张卡、哪张充电网。", "自查单，\n两份书面口径。", "生产、采购、合规"),
        ("02", "带去合作方", "拿三样去谈：\n月十万套的数、\n设备图纸、\n育苗就开始的合同。", "一份按量重算的报价，\n一份栽培合同。", "采购、企划"),
        ("03", "明年年底前", "中文骨架，\n一百到二百支上线。\n首页只摆两样。\n语音只确认这两样。", "骨架清单，\n首页与语音验收单。", "企划、培训、数字"),
    ]
    x = ML
    for i, (num, head, todo, out, who) in enumerate(steps3):
        if i:
            rect(s, x - 0.22, 1.3, 0.012, 4.35, HAIR)
        tb(s, x, 1.3, 3.5, 0.26, [{"text": num, "font": MED, "size": 14, "color": RED}])
        tb(s, x, 1.6, 3.6, 0.5, [{"text": head, "font": MED, "size": 22, "color": INK}])
        tb(s, x, 2.25, 3.5, 0.26, [{"text": "做什么", "font": SANS, "size": 12, "color": RED}])
        tb(s, x, 2.52, 3.6, 1.4, [{"text": todo, "font": SANS, "size": 14, "color": BODY, "line": 1.3}])
        tb(s, x, 4.0, 3.5, 0.26, [{"text": "交出什么", "font": SANS, "size": 12, "color": RED}])
        tb(s, x, 4.27, 3.6, 0.8, [{"text": out, "font": SANS, "size": 14, "color": BODY, "line": 1.3}])
        tb(s, x, 5.05, 3.5, 0.26, [{"text": "建议牵头", "font": SANS, "size": 12, "color": RED}])
        tb(s, x, 5.32, 3.6, 0.4, [{"text": who, "font": SANS, "size": 14, "color": BODY}])
        x += 4.05
    rect(s, ML, 5.95, 11.9, 0.012, HAIR)
    tb(
        s, ML, 6.08, 11.8, 0.7,
        [{"text": "每月看六个数：型号数、换料清洗次数、罐容与人数，\n建了记录的人、连续用的天数、每天打开阿扎的人。", "font": MED, "size": 16, "color": RED, "line": 1.3}],
    )
    folio(s, 14)
    notes(s, "时间点只有一个是会长讲过的：明年年底前的一百到二百支骨架。首页两样和语音确认与骨架同一批验收。其余按先后排，不标日期。牵头部门是建议。")

    # 14 三个坑
    s = new(prs)
    bg(s, WHITE)
    claim(s, "14", "三个坑，中国法人各接一个")
    pits = [
        ("01", "合规", "牌照优先于放量；认证模板复制；收益封顶。",
         "艾护理哪句话能说、记录存哪里，阿扎哪张卡。定之前，不对会员说。"),
        ("02", "供应链与地缘", "多基地生产，本地定价对冲。",
         "原料在中国，下一道工序放在墙外。转运少一趟，风险就少一段。"),
        ("03", "模式声誉", "伦理自律委员会执法，违规清退，透明披露。",
         "一百到二百支骨架把话术固定。事业者换脸，不换话。乱宣传，先从话术上收。"),
    ]
    x = ML
    for i, (num, head, manual, ours) in enumerate(pits):
        if i:
            rect(s, x - 0.22, 1.3, 0.012, 4.1, HAIR)
        tb(s, x, 1.3, 3.5, 0.26, [{"text": num, "font": MED, "size": 14, "color": RED}])
        tb(s, x, 1.6, 3.6, 0.5, [{"text": head, "font": MED, "size": 22, "color": INK}])
        tb(s, x, 2.3, 3.5, 0.26, [{"text": "手册怎么填", "font": SANS, "size": 12, "color": RED}])
        tb(s, x, 2.57, 3.6, 1.0, [{"text": manual, "font": SANS, "size": 14, "color": BODY, "line": 1.3}])
        tb(s, x, 3.75, 3.5, 0.26, [{"text": "中国法人接什么", "font": SANS, "size": 12, "color": RED}])
        tb(s, x, 4.02, 3.6, 1.4, [{"text": ours, "font": SANS, "size": 14, "color": BODY, "line": 1.3}])
        x += 4.05
    rect(s, ML, 5.7, 11.9, 0.012, HAIR)
    tb(s, ML, 5.9, 11.8, 0.9, [{"text": "手册的判断：走砸，多半是合规，\n不是市场不够大。", "font": MED, "size": 20, "color": RED, "line": 1.3}])
    folio(s, 15)
    notes(s, "左边一行来自《养成手册》第九章的填坑指南，右边是把它落到中国法人已经定下的动作上。手册里关于中国牌照状态的表述没有引用，以法务口径为准。")

    # 15 收束
    s = new(prs)
    bg(s, DARK)
    rect(s, 0, 0, 0.08, H, RED)
    lines = [
        ("产线", "烟台四条线，先盘点"),
        ("合作方", "一个数、一张图纸、一份合同"),
        ("企划", "一次写完，一次通到终点"),
        ("三 A", "记录、那一个、每天"),
    ]
    y = 0.72
    for a, b in lines:
        tb(s, 0.85, y, 5.0, 0.5, [{"text": a, "font": MED, "size": 26, "color": WHITE}])
        tb(s, 6.15, y + 0.08, 6.3, 0.4, [{"text": b, "font": SANS, "size": 16, "color": "B5B5B5"}])
        y += 0.78
    tb(s, 0.82, 4.05, 11, 0.85, [{"text": "不能不去", "font": MED, "size": 48, "color": RED}])
    tb(s, 0.88, 5.05, 11.2, 0.85, [{"text": "每月看六个数：型号数、换料清洗次数、罐容与人数，\n建了记录的人、连续用的天数、每天打开阿扎的人。", "font": SANS, "size": 18, "color": "D4D4D4", "line": 1.25}])
    tb(s, 0.88, 6.35, 9, 0.3, [{"text": "蒙想讯息    2026年9月22日", "font": SANS, "size": 14, "color": DIM}])
    folio(s, 16, dark=True)
    notes(s, "收束回到四项：产线盘点、合作方三样、企划、三A，以及每月看的六个数。然后停住。")

    return prs


if __name__ == "__main__":
    path = "/workspace/艾多美何以成为100兆韩元企业.pptx"
    build().save(path)
    print(path)
