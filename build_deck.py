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
    total_note = "全稿 8 页"

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
        [{"text": "APP 是事业者自己的店。Acare 对着这个人说缺什么。\nAZA 把每天的开支接进来。写给中国法人。", "font": SANS, "size": 18, "color": "D4D4D4", "line": 1.4}],
    )
    folio(s, 1, dark=True)
    notes(s, total_note + "。只讲 3A：APP、Acare、AZA，以及人工智能怎样抬体验、抬业绩。")

    # 02 一条线
    s = new(prs)
    bg(s, WHITE)
    claim(s, "01", "3A 是一条线")
    chain = [
        ("01", "Acare 建记录", "新人先进来。\n每餐、体型、步数、\n用药。\n记录先每天跑起来。"),
        ("02", "说缺什么", "先说这个人\n缺哪个功能。\n一个功能，\n只有一个产品。"),
        ("03", "那一个", "进的就是品类里\n那一个冠军。\n不必再比较，\n话对着这个人说。"),
        ("04", "AZA 每天", "每天的开支接进来。\n人每天打开。\n主力品目在这里\n被带出来。"),
        ("05", "记在谁名下", "买的人记在这个\n平台主人名下。\n连续用了几天，\n每天打不打开，\n都从这里看见。"),
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
        [{"text": "3A 就是 APP、Acare、AZA。\n要数三样：建了记录的人，连续用的天数，每天打开 AZA 的人。", "font": MED, "size": 20, "color": RED, "line": 1.3}],
    )
    folio(s, 2)
    notes(s, "3A 就是 APP、Acare、AZA，不是三个并列的应用。新人进 Acare 建记录，记录说缺哪个功能，一能一品定那一个产品，AZA 每天接开支带出主力品目，买的人记在平台主人名下。平台主人是公司里的 PPO。要数的三样由中国法人自己给出，这里不填数字。")

    # 03 APP
    s = new(prs)
    bg(s, WHITE)
    claim(s, "02", "APP，先改变事业者")
    app = [
        ("01", "八个阶段在里面", "产品懂得多少，\n八个阶段走到哪一步，\n下一步学什么，都在里面。\n中国法人做成中文的路径。\n人不用等一场酒店课。"),
        ("02", "一百到二百支骨架", "到明年年底先打好骨架。\n事业者换成自己的脸、\n自己的讲法。\n一小时收成五分钟。\n说明在这个人的平台上完成。"),
        ("03", "平台主人", "事业者是这家平台的主人，\n不是发链接的人。\n公司里的叫法是 PPO。\n店、健康、教育，收在一处。\n培育平台主人，是九月十三日\n三份发表之一。"),
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
        [{"text": "酒店课的费用在这里消失。\n省下来的，补在事业者那三十五个点上。", "font": MED, "size": 22, "color": RED, "line": 1.3}],
    )
    folio(s, 3)
    notes(s, "APP 先武装事业者：阶段、中文骨架、平台主人。PPO 是公司里对个人平台主人的叫法，全称 Personal Platform Owner，2021 年会刊已用。培育平台主人是9月13日奖项的三份发表之一。酒店课是旧成本。")

    # 04 Acare
    s = new(prs)
    bg(s, WHITE)
    claim(s, "03", "Acare 对着这个人说")
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
        ("先进入 Acare", "介绍别人，先把对方的记录建起来，再谈产品。"),
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
    folio(s, 4)
    notes(s, "Acare 的输入是每餐、体型、步数、用药。输出先是这个人缺哪个功能，产品因一能一品而跟着定，人工智能不是在几个产品里挑。血和尿可以后补，记录不能等医院。记录存哪里、哪句话能说，中国要先定，定之前不对会员说。")

    # 05 AZA
    s = new(prs)
    bg(s, WHITE)
    claim(s, "04", "AZA 让人每天走进来")
    aza = [
        ("01", "十万个，不赚一百兆", "鸡蛋、油、充电接进来。\n一百兆由大约五百个\n主力品目赚。\n这十万个负责把人\n每天引进来。"),
        ("02", "鲭鱼是个例子", "会长拿鲭鱼作例子。\n这种日常品成本率\n到七成五。单靠它\n到不了一百兆。\n它让人每天来。"),
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
    folio(s, 5)
    notes(s, "AZA 是市集的地面，不是行政窗口。鲭鱼出自9月22日文稿里的反问：成本率 75% 的日常品，35% 从哪来。答案是单靠它到不了一百兆，它让人每天来，五百个主力品目负责赚。充电接不接、走哪张卡哪张网，中国先定，不卖加油券。")

    # 06 商城
    s = new(prs)
    bg(s, WHITE)
    claim(s, "05", "首页两样，语音只确认")
    mall = [
        ("01", "千人千面", "打开不是目录。今天只摆两样：缺的那个功能对应的产品，旁边一句为什么；再加上每天用的那一件。人不同，两样不同。功效句只用审过的。没定哪句话能说，首页不新写。"),
        ("02", "推荐只做三件", "缺什么：记录决定功能，功能决定那一个，不另推替代款。快用完了：按上次购买和用量提醒续上。该带给谁：先发进入 Acare 的入口，不发产品链接。"),
        ("03", "语音加购与支付", "说“把今天这两样加入购物车”，再说“确认支付”。语音不搜索、不比较、不改款。支付走现有收银台，语音里不报卡号、不存卡号。收银台没接上之前，语音只加购。"),
        ("04", "下完单，记在谁名下", "记在这个平台主人名下。这个人的连续使用更新。事业者工作台只看三行：谁还没建记录，谁的瓶子快空了，谁停了几天。"),
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
    folio(s, 6)
    notes(s, "千人千面是两样和那句话因人而异，不是目录变长。推荐只做缺什么、快用完、该带给谁。语音只确认已经摆上的两样，支付走现有收银台，不在语音里收集卡号。下单记在平台主人名下，并更新连续使用。")

    # 07 四件延展
    s = new(prs)
    bg(s, WHITE)
    claim(s, "06", "四件都长在两样和语音上")
    more = [
        ("01", "拍完，首页当场重算", "拍下这一餐，两样立刻换，不等运营改图。当次就能加购，不流失到下次打开。"),
        ("02", "空瓶前三天，只续同一款", "按上次的量和每天用量倒推。提前三天，语音只问要不要续上。不推新品。连续天数不断在这里。"),
        ("03", "家里的人分开记", "父母、孩子各一份记录。首页按人出两样，不混进一个购物车。买对的人变多，一户不止一个人在用。"),
        ("04", "停了才开口，而且分档", "三天、七天、十四天，话用审过的骨架，对象是这一个人。点开就联系。被记得，不是被群发。"),
    ]
    y = 1.28
    for num, head, body in more:
        tb(s, ML, y, 0.7, 0.36, [{"text": num, "font": MED, "size": 16, "color": RED}])
        tb(s, 1.55, y - 0.02, 10.9, 0.4, [{"text": head, "font": MED, "size": 22, "color": INK}])
        tb(s, 1.55, y + 0.42, 10.9, 0.55, [{"text": body, "font": SANS, "size": 16, "color": BODY}])
        y += 1.15
    rect(s, ML, 5.9, 11.9, 0.012, HAIR)
    tb(
        s, ML, 6.08, 11.8, 0.7,
        [{"text": "体验省掉的是改图、再找、买错、被群发。业绩留下的是当次加购、续上、一户多人、停掉的人回来。", "font": MED, "size": 16, "color": RED}],
    )
    folio(s, 7)
    notes(s, "四件都接在已有的首页两样和语音确认上，不另开一个会在几百款里挑的引擎。拍完重算是当次转化。空瓶前三天只续同一款。家里的人分开记，首页按人出。停用分三天、七天、十四天，用审过的骨架对这一个人说。")

    # 08 收束
    s = new(prs)
    bg(s, DARK)
    rect(s, 0, 0, 0.08, H, RED)
    lines = [
        ("APP", "五分钟说明，早上三行"),
        ("Acare", "拍一张，说缺什么"),
        ("AZA", "说一句，就买到"),
        ("要数的", "记录、连续天数、每天打开"),
    ]
    y = 0.72
    for a, b in lines:
        tb(s, 0.85, y, 5.0, 0.5, [{"text": a, "font": MED, "size": 26, "color": WHITE}])
        tb(s, 6.15, y + 0.08, 6.3, 0.4, [{"text": b, "font": SANS, "size": 16, "color": "B5B5B5"}])
        y += 0.78
    tb(s, 0.82, 4.05, 11, 0.85, [{"text": "不能不去", "font": MED, "size": 48, "color": RED}])
    tb(s, 0.88, 5.05, 11.2, 0.85, [{"text": "体验省掉的是等课、填表、搜索。\n业绩留下的是续购，和跟到的那一个人。", "font": SANS, "size": 18, "color": "D4D4D4", "line": 1.25}])
    tb(s, 0.88, 6.35, 9, 0.3, [{"text": "蒙想讯息    2026年9月22日", "font": SANS, "size": 14, "color": DIM}])
    folio(s, 8, dark=True)
    notes(s, "收束只留 3A：APP、Acare、AZA，以及要数的三样。然后停住。")

    return prs


if __name__ == "__main__":
    path = "/workspace/艾多美何以成为100兆韩元企业.pptx"
    build().save(path)
    print(path)
