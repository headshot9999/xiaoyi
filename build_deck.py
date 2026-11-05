#!/usr/bin/env python3
"""艾多美何以成为100兆韩元企业 —— 参照 CN_DAX 课题文稿重新构建演示文稿。"""

from pathlib import Path
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SANS = "Noto Sans CJK SC"
MED = "Noto Sans CJK SC Medium"

WHITE = "FFFFFF"
OFF_WHITE = "F7F7F6"
CARD_BG = "F2F1EF"
DARK_CARD = "1C1C1F"
INK = "141414"
BODY = "2C2C2C"
MUTED = "6A6A6A"
HAIR = "E2E2E2"
RED = "9E1B1B"
RED_SOFT = "F9EEEE"
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


def rect(slide, x, y, w, h, fill, line_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    if line_color:
        shape.line.color.rgb = rgb(line_color)
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    return shape


def bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def new(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def folio(slide, n, total=10, dark=False):
    tb(
        slide, 11.2, 7.08, 1.45, 0.24,
        [{"text": f"{n:02d} / {total:02d}", "font": SANS, "size": 10.5, "color": "6A6A6A" if dark else "B0B0B0", "align": "right"}],
    )


def claim(slide, num, tag, text):
    tb(slide, ML, 0.34, 9.5, 0.28, [{"text": f"{num}  ·  {tag}", "font": MED, "size": 12.5, "color": RED}])
    tb(slide, ML, 0.66, 11.9, 0.48, [{"text": text, "font": MED, "size": 26.5, "color": INK}])
    rect(slide, ML, 1.24, 11.9, 0.014, HAIR)


def footer_banner(slide, text):
    rect(slide, ML, 5.82, 11.9, 0.92, RED_SOFT)
    rect(slide, ML, 5.82, 0.07, 0.92, RED)
    tb(
        slide, ML + 0.26, 5.82, 11.45, 0.92,
        [{"text": text, "font": MED, "size": 16.5, "color": RED, "line": 1.32}],
        valign="middle",
    )


def build():
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    prs.core_properties.title = "艾多美何以成为100兆韩元企业"
    prs.core_properties.subject = "从直视危机到结构转型 · 中国法人 DAX 落地实践"
    total_pages = 10

    # =========================================================================
    # 01 封面
    # =========================================================================
    s = new(prs)
    bg(s, DARK)
    rect(s, 0, 0, 0.09, H, RED)
    tb(s, 0.88, 1.05, 10, 0.3, [{"text": "100兆课题论证  ·  中国法人 DAX（Digital & AI Transformation）", "font": MED, "size": 13, "color": "C87A7A"}])
    tb(
        s, 0.85, 1.72, 11.5, 2.1,
        [
            {"text": "艾多美何以必然成为", "font": MED, "size": 46, "color": WHITE, "after": 4},
            {"text": "100兆韩元企业", "font": MED, "size": 46, "color": WHITE},
        ],
    )
    rect(s, 0.88, 4.02, 1.2, 0.04, RED)
    tb(
        s, 0.88, 4.28, 11.2, 1.15,
        [
            {"text": "这不是玫瑰色的乐观口号，而是“如果像现在这样安逸下去真的会死”的生存战略。", "font": MED, "size": 19, "color": "EAEAEA", "after": 8},
            {"text": "从一贯工程、3A（APP · Acare · AZA）到中国电算团队率先填补的数据与 AI 空白。", "font": SANS, "size": 15.5, "color": "A8A8A8"},
        ],
    )
    pillars = [
        ("01  直视危机与必要条件", "3% 增速警报 · 破除“企业麻风病” · 35% 返还安全网"),
        ("02  三大增长引擎", "一贯工程（先胜求战） · 3A 平台（PPO） · 三无与上限线"),
        ("03  中国 DAX 率先填空", "阿里云数仓（PIPL 合规） · AI 编程内化 · 跨部门 AI 智能体"),
    ]
    px = 0.88
    for title_p, desc_p in pillars:
        rect(s, px, 5.65, 3.72, 1.08, DARK_CARD)
        rect(s, px, 5.65, 3.72, 0.03, RED)
        tb(s, px + 0.18, 5.8, 3.36, 0.32, [{"text": title_p, "font": MED, "size": 13, "color": WHITE}])
        tb(s, px + 0.18, 6.16, 3.36, 0.45, [{"text": desc_p, "font": SANS, "size": 11.5, "color": DIM, "line": 1.25}])
        px += 3.92
    folio(s, 1, total_pages, dark=True)
    notes(s, "全稿10页。从直视危机与“企业麻风病”出发，完整阐述为何100兆是生存必要条件、一贯工程先胜求战、3A（APP·Acare·AZA）平台、三无与Chutzpah精神，以及中国电算团队（DAX）率先填补的三项空白。")

    # =========================================================================
    # 02 序言：从直视危机开始
    # =========================================================================
    s = new(prs)
    bg(s, WHITE)
    claim(s, "01", "序言 · 直视危机", "“如果像现在这样安逸下去，真的会死”——结构转型的最后黄金时间")

    rect(s, ML, 1.48, 3.65, 4.05, DARK)
    tb(s, ML + 0.28, 1.72, 3.1, 0.3, [{"text": "过去 10 年年均增速", "font": SANS, "size": 13, "color": DIM}])
    tb(s, ML + 0.28, 2.02, 3.1, 0.72, [{"text": "20% ~ 30%", "font": MED, "size": 36, "color": WHITE}])
    rect(s, ML + 0.28, 2.88, 3.05, 0.015, "333333")
    tb(s, ML + 0.28, 3.08, 3.1, 0.3, [{"text": "去年实际增速（擦线过关）", "font": SANS, "size": 13, "color": "E09090"}])
    tb(s, ML + 0.28, 3.38, 3.1, 0.85, [{"text": "仅 3%", "font": MED, "size": 44, "color": RED}])
    tb(
        s, ML + 0.28, 4.35, 3.1, 0.95,
        [{"text": "在网络营销结构中，增长停滞意味着事业者奖金缩水、信任动摇，组织基础可能瞬间崩塌。", "font": SANS, "size": 13, "color": "CCCCCC", "line": 1.32}],
    )

    crises = [
        ("01  外部强敌蚕食", "“先把价格钉死、再去配品质”的对手逼近", "像大创（Daiso）这样先钉死价格再倒推品质的强悍竞争者正在快速蚕食大众消费市场，单纯靠过去的惯性已无法守住阵地。"),
        ("02  内部“企业麻风病”", "核心合作社破产重整、供应方连环亏损却毫无察觉", "朴韩吉会长将其比作感觉不到疼痛而任由身体腐烂的“企业麻风病（기업적 문둥병）”——这是慢慢杀死企业最致命的病。"),
        ("03  最后的黄金时间", "100兆不是沿袭旧经验，而是把整盘棋推倒重来", "沿用过去的成功方式绝不可能到达100兆。今年被钉死为“结构转型的最后黄金时间”，必须彻底砸碎旧框、全面换盘。"),
    ]
    cy = 1.48
    for c_num, c_head, c_body in crises:
        rect(s, 4.62, cy, 8.0, 1.24, OFF_WHITE, HAIR)
        tb(s, 4.86, cy + 0.14, 2.6, 0.26, [{"text": c_num, "font": MED, "size": 12, "color": RED}])
        tb(s, 4.86, cy + 0.40, 7.5, 0.34, [{"text": c_head, "font": MED, "size": 17, "color": INK}])
        tb(s, 4.86, cy + 0.76, 7.5, 0.42, [{"text": c_body, "font": SANS, "size": 13, "color": BODY, "line": 1.25}])
        cy += 1.40

    footer_banner(s, "结论：100兆不是玫瑰色的口号，而是从凛冽危机感出发的生存宣言——今年是结构转型的最后黄金时间。")
    folio(s, 2, total_pages)
    notes(s, "过去10年年均20~30%增长，去年仅3%。外部有大创等先定价再配品质的对手，内部有合作社法定管理却浑然不知的“企业麻风病”。")

    # =========================================================================
    # 03 第一章：不是“想变大”，而是“不变大就活不下去”
    # =========================================================================
    s = new(prs)
    bg(s, WHITE)
    claim(s, "02", "第一章 · 生存的必要条件", "不是“想变大”，而是“不变大就活不下去”：35% 返还是真正的社会安全网")

    cols = [
        (
            "01  无就业增长已是现实",
            "AI 正在抹去白领与专业岗",
            "AI 不仅替代产线，更在抹去办公室白领与专业职位。韩国前十大企业一年间蒸发数万个岗位。\n\n但即便工作消失，人依然要吃、要涂、要洗——消费绝不会停止。",
        ),
        (
            "02  35% 返还胜过“机器人税”",
            "支撑普通人养老的社会安全网",
            "把过去被大型流通财阀和大股东独占的约 35% 流通毛利，原封不动返还给消费者。\n\n比起学者停留在口头上的“机器人税”，这套“35% Refund”才是最现实的养老安全网。",
        ),
        (
            "03  体量决定能否托底",
            "从 10兆、100兆到下一代 1000兆",
            "停留在国内第一的 10兆规模，只能造就少数顶层领袖，托不起普通大多数人的养老。\n\n必须把流向沃尔玛、亚马逊、大型商超的消费拉进来：会长有生之年（约15年）做到100兆，下一代迈向1000兆资金池。",
        ),
    ]
    cx = ML
    for c_tag, c_title, c_desc in cols:
        rect(s, cx, 1.48, 3.76, 4.08, OFF_WHITE, HAIR)
        rect(s, cx, 1.48, 3.76, 0.05, RED)
        tb(s, cx + 0.24, 1.70, 3.28, 0.28, [{"text": c_tag, "font": MED, "size": 12.5, "color": RED}])
        tb(s, cx + 0.24, 2.04, 3.28, 0.45, [{"text": c_title, "font": MED, "size": 19.5, "color": INK}])
        rect(s, cx + 0.24, 2.58, 3.28, 0.012, HAIR)
        tb(s, cx + 0.24, 2.74, 3.28, 2.6, [{"text": c_desc, "font": SANS, "size": 14, "color": BODY, "line": 1.38}])
        cx += 4.07

    footer_banner(s, "100兆的本质不是单纯的野心，而是让普通大多数人跨过“岗位消失浪潮”所必需的资金池体量。")
    folio(s, 3, total_pages)
    notes(s, "AI抹去岗位，但消费不停。35%流通毛利返还是比机器人税更现实的社会安全网。10兆只够少数领袖，100兆（15年）至1000兆才能托住大众养老。")

    # =========================================================================
    # 04 第二章：削去赘肉的“一贯工程”，与半价决胜手
    # =========================================================================
    s = new(prs)
    bg(s, WHITE)
    claim(s, "03", "第二章 · 一贯工程与半价决胜", "一能一品 · 先胜求战（先钉死必胜价格，再倒推工序重做成本）")

    # Left column: 一能一品 & 70% 大众
    rect(s, ML, 1.48, 5.75, 4.08, OFF_WHITE, HAIR)
    tb(s, ML + 0.26, 1.66, 5.2, 0.3, [{"text": "01  一能一品（1功能1产品）与 70% 大众市场", "font": MED, "size": 14, "color": RED}])
    left_items = [
        ("终结多品种小批量", "把 4~5 款洗发水整合成压倒性品质的“唯一 1 款”，彻底砍掉分摊成本与库存负担。"),
        ("精准锁定 70% 大众", "放弃只买昂贵品的顶层 1% 与纯低价市场，向追求合理消费的 70% 大众以半价提供名品。"),
        ("容量翻倍重置已启动", "洗发水、沐浴露同价从 500ml 直接翻倍至 1,000ml；目标是每个进入品目拿下全球 50% 份额。"),
    ]
    ly = 2.08
    for l_h, l_b in left_items:
        tb(s, ML + 0.26, ly, 5.2, 0.3, [{"text": f"▪  {l_h}", "font": MED, "size": 15, "color": INK}])
        tb(s, ML + 0.48, ly + 0.30, 5.0, 0.62, [{"text": l_b, "font": SANS, "size": 13, "color": BODY, "line": 1.28}])
        ly += 1.08

    # Right column: 先胜求战 & 实证
    rect(s, 6.72, 1.48, 5.9, 2.22, OFF_WHITE, HAIR)
    tb(s, 6.96, 1.66, 5.4, 0.3, [{"text": "02  先胜求战：不压合作社单价，而是共同拆改工序", "font": MED, "size": 14, "color": RED}])
    tb(
        s, 6.96, 2.04, 5.4, 1.5,
        [
            {"text": "1. 果断砍掉只抬高价格的营销噱头附加成分", "font": SANS, "size": 13.5, "color": INK, "after": 6},
            {"text": "2. 将容器（包材）占成本的比重压低至 1/4 水平", "font": SANS, "size": 13.5, "color": INK, "after": 6},
            {"text": "3. 以 5 倍、10 倍以上销量为前提，搭建大批量自动化产线", "font": SANS, "size": 13.5, "color": INK},
        ],
    )

    rect(s, 6.72, 3.86, 5.9, 1.70, DARK)
    tb(s, 6.96, 4.02, 5.4, 0.28, [{"text": "03  已被验证的“十分之一价格 + 压倒性销量”语法", "font": MED, "size": 13, "color": "E09090"}])
    tb(
        s, 6.96, 4.38, 5.4, 1.05,
        [
            {"text": "▪ HemoHIM：77万韩元 → 10吨自动化设备改造降至 7万韩元档，现月销 30万盒", "font": SANS, "size": 13, "color": WHITE, "after": 8},
            {"text": "▪ 功能性牙刷：市售 2~3千韩元 → 模具固定化量产降至 1,000韩元，累计售出数千万支", "font": SANS, "size": 13, "color": WHITE},
        ],
    )

    footer_banner(s, "从原料、制造、包装、物流到营销绑成一条履带：把“十分之一价格 + 压倒性销量”放大到全品目，就是100兆的实体。")
    folio(s, 4, total_pages)
    notes(s, "一能一品整合4~5款洗发水为1款，同价500ml升至1000ml，瞄准70%大众，目标全球份额50%。先胜求战三招：减噱头成分、容器成本压至1/4、5~10倍销量自动化产线。")

    # =========================================================================
    # 05 第三章（上）：手心里自动运转的平台 3A
    # =========================================================================
    s = new(prs)
    bg(s, WHITE)
    claim(s, "04", "第三章 · 数字化创新 3A", "手心里自动运转的平台：APP · Acare · AZA 构成的 AI 商业引擎")

    cards_3a = [
        (
            "APP",
            "个人平台  ·  AI 商业引擎与 PPO",
            [
                "不是商品罗列式商城，而是在后台替事业者自动运转业务的 AI 商业引擎。",
                "总部将全功能电商免费分发给个人，事业者成为 PPO（个人平台主人）；10年后“网络营销”一词将被“PPO”取代。",
                "过去商城像抽号排队的银行窗口，未来 APP 是热闹市集：AI 主动搭话，一两句日常对话即可完成会员注册与推荐。",
                "教育彻底重构：总部高清标准视频一键全球共享；本人头像合成为 AI 数字人（Avatar），用多国语言讲解业务。",
            ],
        ),
        (
            "Acare",
            "艾护理  ·  免费健康管理锁客（Lock-in）",
            [
                "作为牢牢锁住会员的免费健康管理装置，打通身高、体重、饮食拍照、运动与健康测量数据。",
                "不再对所有人背同一套推销话术，而是由 AI 洞察个体身体数据与生活习惯，指出缺哪个功能。",
                "结合“一能一品”原则，缺哪个功能就直接对应那一个品类冠军产品，让选择成本归零。",
                "数据越连续，判断越准，会员越没有理由离开平台。",
            ],
        ),
        (
            "AZA",
            "阿扎  ·  高频日常流通与威海物流枢纽",
            [
                "跨越传统日用品，向一般流通商品与 K-Food 大幅扩展，把每日生活开支接进同一个平台。",
                "十万级日常品目负责像市集一样让人每天走进来，带出主力品目的持续复购。",
                "战略落子中国山东半岛威海（Weihai）物流枢纽，同时拿下增值税退税与成本节降双重优势。",
                "人每天因为高频消费打开平台，主力品目的超大规模产量才留得住。",
            ],
        ),
    ]
    ax = ML
    for a_code, a_sub, a_bullets in cards_3a:
        rect(s, ax, 1.48, 3.76, 4.12, OFF_WHITE, HAIR)
        rect(s, ax, 1.48, 3.76, 0.52, DARK)
        tb(s, ax + 0.20, 1.56, 1.2, 0.36, [{"text": a_code, "font": MED, "size": 18, "color": RED}])
        tb(s, ax + 1.15, 1.60, 2.45, 0.32, [{"text": a_sub, "font": SANS, "size": 11.5, "color": WHITE, "align": "right"}])
        by = 2.14
        for b_text in a_bullets:
            tb(s, ax + 0.20, by, 3.36, 0.80, [{"text": f"• {b_text}", "font": SANS, "size": 12.8, "color": BODY, "line": 1.26}])
            by += 0.84
        ax += 4.07

    footer_banner(s, "3A 不是三个并列的应用，而是一条自动运转的链条：Acare 锁住健康关系，AZA 拓宽日常消费触点，APP 让 PPO 事业自动运转。")
    folio(s, 5, total_pages)
    notes(s, "APP是后台自动运转的AI商业引擎，事业者成为PPO，本人头像合成AI数字人多语言讲解；Acare免费健康管理锁客；AZA扩向一般流通与K-Food，并在山东威海设物流枢纽实现增值税退税与降本。")

    # =========================================================================
    # 06 第三章（下）：3A 在商城与展业中的 AI 体验落地
    # =========================================================================
    s = new(prs)
    bg(s, WHITE)
    claim(s, "05", "第三章延展 · 3A 体验与业绩转化", "从“抽号排队的银行窗口”到“AI 先开口搭话的热闹市集”")

    flows = [
        ("01", "对话式入会与千人千面首页", "会员注册在一两句日常对话中完成。打开商城不再是冗长目录，AI 根据身高、体重、膳食拍照与购买习惯主动搭话，首页当场重算，只摆今天最需要的两样（精准功能品 + 日常高频品）。"),
        ("02", "空瓶前三天预警与语音下单", "按上次购买量与每日用量倒推，空瓶前三天主动提醒续上同一款；用户直接说“加入购物车、确认支付”，走现有合规收银台完成下单，让连续使用天数不断档。"),
        ("03", "家庭成员分档管理", "父母、伴侣、孩子在 Acare 各建独立健康档案，首页按家庭成员分别生成定制方案，不混入同一推荐池，推动“一户多人使用”的业绩倍增。"),
        ("04", "PPO 工作台自动跟进与数字人分身", "工作台每天只推三行关键名单：谁还没建档案、谁的瓶子快空了、谁停用了3/7/14天。点开即可发送本人头像合成的 5 分钟标准 AI 数字人视频，被记得而不是被群发。"),
    ]
    fy = 1.44
    for f_num, f_head, f_body in flows:
        rect(s, ML, fy, 11.9, 0.98, OFF_WHITE, HAIR)
        tb(s, ML + 0.22, fy + 0.16, 0.6, 0.32, [{"text": f_num, "font": MED, "size": 15, "color": RED}])
        tb(s, ML + 0.82, fy + 0.14, 3.1, 0.68, [{"text": f_head, "font": MED, "size": 16.5, "color": INK, "line": 1.25}])
        tb(s, ML + 4.05, fy + 0.14, 7.6, 0.72, [{"text": f_body, "font": SANS, "size": 13.5, "color": BODY, "line": 1.28}])
        fy += 1.08

    footer_banner(s, "用户体验省掉的是填表、搜索、比价与等课；业绩增长来自当次转化、不断档续购、一户多人与精准唤醒。")
    folio(s, 6, total_pages)
    notes(s, "把文稿中“AI先开口搭话、对话完成注册、本人头像合成视频”与3A商城千人千面、语音加购、PPO工作台结合，具体展示AI如何抬升用户体验与业绩。")

    # =========================================================================
    # 07 第四章：奖金的上限线，与现场的“Chutzpah”精神
    # =========================================================================
    s = new(prs)
    bg(s, WHITE)
    claim(s, "06", "第四章 · 制度根基与组织姿态", "奖金的“三无与上限线”，以及现场的“Chutzpah（死磕破局）”精神")

    # Left: 奖金制度
    rect(s, ML, 1.48, 5.75, 4.10, OFF_WHITE, HAIR)
    rect(s, ML, 1.48, 5.75, 0.05, RED)
    tb(s, ML + 0.26, 1.68, 5.2, 0.32, [{"text": "01  吸引人的制度根基：“三无（3無）”与上限线", "font": MED, "size": 15, "color": RED}])
    tb(
        s, ML + 0.26, 2.12, 5.2, 3.2,
        [
            {"text": "▪ 无维持条件、无层级限制、设有奖金上限线（封顶）。", "font": MED, "size": 14.5, "color": INK, "after": 10},
            {"text": "▪ 无限阶段的力量：他社层级截断，必须直推精英才能成功；艾多美第 100 层、1000 层下爆发的业绩也完整计入我的销售额。", "font": SANS, "size": 13.5, "color": BODY, "line": 1.32, "after": 10},
            {"text": "▪ 跨国奇迹的制度来源：韩国小学学历事业者培育出台湾领袖，再延伸出中国最高职级者，正源于无限阶段。", "font": SANS, "size": 13.5, "color": BODY, "line": 1.32, "after": 10},
            {"text": "▪ 封顶守住底层：不让一个人独拿月薪 10 亿韩元，而是让十个人各拿月薪 1 亿韩元，确保底层稳固不崩塌。", "font": SANS, "size": 13.5, "color": BODY, "line": 1.32},
        ],
    )

    # Right: 内部员工与事业者姿态
    rect(s, 6.72, 1.48, 5.9, 4.10, OFF_WHITE, HAIR)
    rect(s, 6.72, 1.48, 5.9, 0.05, DARK)
    tb(s, 6.98, 1.68, 5.4, 0.32, [{"text": "02  推动系统的人：第二次创业、Chutzpah 与婴儿哲学", "font": MED, "size": 15, "color": INK}])
    tb(
        s, 6.98, 2.12, 5.4, 3.2,
        [
            {"text": "▪ 第二次创业成员：员工不是坐办公室抱签呈板的管理员，必须深入合作社产线最底层，一起拆改成本结构。", "font": SANS, "size": 13.5, "color": BODY, "line": 1.32, "after": 10},
            {"text": "▪ 拒绝“野鸡埋雪”式部门利己主义：会长把“只管好自己部门就行”比作把头埋进雪里以为安全的野鸡。", "font": SANS, "size": 13.5, "color": BODY, "line": 1.32, "after": 10},
            {"text": "▪ Chutzpah 精神与“有对策的越界”：拿出赌上职位也要激烈碰撞、说服到底的死磕精神（후츠파）与跨界担当。", "font": MED, "size": 13.5, "color": RED, "line": 1.32, "after": 10},
            {"text": "▪ 婴儿哲学 vs 奶牛哲学：事业者要把伙伴当作需要用爱照料的“婴儿”，而非当成产奶工具的“奶牛”，组织才长久。", "font": SANS, "size": 13.5, "color": BODY, "line": 1.32},
        ],
    )

    footer_banner(s, "制度再好，如果推动它的人停在原地也只是空谈：用 Chutzpah 精神打破部门高墙，用婴儿哲学培育伙伴。")
    folio(s, 7, total_pages)
    notes(s, "三无（无维持、无层级、有上限线）：100/1000层业绩全计入，10人各拿1亿而非1人拿10亿。员工要有第二次创业觉悟，抛弃“野鸡把头埋进雪里”的本位主义，拿出赌上职位的Chutzpah精神与有对策的越界；事业者秉持婴儿哲学而非奶牛哲学。")

    # =========================================================================
    # 08 第五章（上）：100兆不是预言，而是我们要填上的计算表
    # =========================================================================
    s = new(prs)
    bg(s, WHITE)
    claim(s, "07", "第五章 · 使命与现实落差", "100兆不是躺在柿子树下等掉下来的柿子，而是必须填上的计算表")

    # Top metrics row
    metrics = [
        ("1.9兆 → 1.5兆+", "2024年创造约1.9兆韩元经济价值中，1.5兆以上分配给员工、合作社、政府与社区——底层引擎是“解决身边人养老”的使命。"),
        ("2兆  ·  27个法人", "当前握在手里的冷静现实：年销售额约 2兆韩元、全球 27个海外法人、15年时间线。距离 100兆仍有巨大落差。"),
    ]
    mx = ML
    for m_val, m_desc in metrics:
        rect(s, mx, 1.46, 5.8, 1.38, DARK)
        tb(s, mx + 0.26, 1.60, 5.3, 0.42, [{"text": m_val, "font": MED, "size": 24, "color": RED}])
        tb(s, mx + 0.26, 2.08, 5.3, 0.66, [{"text": m_desc, "font": SANS, "size": 12.5, "color": "DDDDDD", "line": 1.28}])
        mx += 6.1

    tb(s, ML, 3.04, 11.9, 0.34, [{"text": "横亘在 2兆 与 100兆 之间、必须用实绩填上的三道空白：", "font": MED, "size": 16, "color": INK}])

    blanks = [
        ("空白 01", "AZA 商城的体量跃升", "从目前仍在爬坡的阶段，真正做出有意义的交易体量增长，把高频日常消费规模化锁进平台。"),
        ("空白 02", "35% 返还的精密财务方程", "在产品价格砍半、免费提供 Acare 健康护理的同时，依然能毫不动摇地兑现 35% 返还的精密财务模型。"),
        ("空白 03", "每季度必须证明的具体数字", "各海外法人与各部门不能只停留在战略口号，必须按季度交出可核验、可累加的具体业务指标。"),
    ]
    bx = ML
    for b_num, b_title, b_body in blanks:
        rect(s, bx, 3.48, 3.76, 2.12, OFF_WHITE, HAIR)
        tb(s, bx + 0.22, 3.64, 3.3, 0.26, [{"text": b_num, "font": MED, "size": 12, "color": RED}])
        tb(s, bx + 0.22, 3.94, 3.3, 0.38, [{"text": b_title, "font": MED, "size": 17.5, "color": INK}])
        tb(s, bx + 0.22, 4.40, 3.3, 1.05, [{"text": b_body, "font": SANS, "size": 13.2, "color": BODY, "line": 1.32}])
        bx += 4.07

    footer_banner(s, "靠恐惧驱动的组织很快疲惫，只追逐金钱的组织会在更大利益前背叛；唯有使命感与填满空白的计算表能走完15年。")
    folio(s, 8, total_pages)
    notes(s, "2024年1.9兆经济价值中1.5兆以上回馈社会。面对年销2兆、27个法人、15年时间线，必须填上AZA体量、半价+免费护理下兑现35%返还的财务方程、各法人每季度具体数字三道空白。")

    # =========================================================================
    # 09 第五章（下）：中国电算团队（DAX）正在率先填补的空白
    # =========================================================================
    s = new(prs)
    bg(s, WHITE)
    claim(s, "08", "第五章落地 · 中国电算团队（DAX）", "IT 部门已改组为 DAX：用数据与 AI 为一贯工程和 3A 铺好底座")

    rect(s, ML, 1.44, 11.9, 0.72, DARK)
    tb(
        s, ML + 0.24, 1.44, 11.4, 0.72,
        [{"text": "组织已完成从 IT 到 DAX（Digital & AI Transformation）的改组  ·  课题不是把独立运行的中国系统生硬合并，而是先对齐数据结构、打好可与总部共享 AI 引擎的底座。", "font": MED, "size": 13.5, "color": WHITE}],
        valign="middle",
    )

    dax_actions = [
        (
            "第一项填空  ·  数据底座与合规",
            "把分散的数据归集到一处",
            "阿里云数仓 + AI BI + PIPL 合规",
            "APP 要记住会员体型、健康数据与购买履历并主动搭话，数据分散在各系统里就无法起步。\n\n我们在阿里云（Alibaba Cloud）上搭建数据仓库，并在其上搭载结合 AI 的内部 BI；按报表切分查阅权限并留存查阅日志，在中国《个人信息保护法》（PIPL）框架内安全用数。",
        ),
        (
            "第二项填空  ·  研发内化与提速",
            "大幅降低对外包的依赖",
            "引入 AI Coding Agent 收回核心开发",
            "3A 一直被指出的弱点正是内部开发能力不足。\n\n中国电算团队将 AI 编程智能体（AI Coding Agent）引入实务，把过去交给外部厂商的开发最大限度收回团队内部。开发速度更快、成本更低，更重要的是建立起“我们自己的系统，我们自己最懂”的结构。",
        ),
        (
            "第三项填空  ·  技术践行跨界担当",
            "用 AI 智能体支援其他部门",
            "技术落实“有对策的越界”",
            "正在设计由 AI Agent 支援各业务部门重复性工作的体系。\n\n会长强调的“有对策的越界（대책 있는 오지랖）”，电算团队正在用技术付诸实践；这与总部在各部门培育 AI Change Agent 的方向完全同频。",
        ),
    ]
    dx = ML
    for d_tag, d_title, d_sub, d_body in dax_actions:
        rect(s, dx, 2.32, 3.76, 3.32, OFF_WHITE, HAIR)
        rect(s, dx, 2.32, 3.76, 0.05, RED)
        tb(s, dx + 0.22, 2.50, 3.32, 0.26, [{"text": d_tag, "font": MED, "size": 11.5, "color": RED}])
        tb(s, dx + 0.22, 2.80, 3.32, 0.36, [{"text": d_title, "font": MED, "size": 18, "color": INK}])
        tb(s, dx + 0.22, 3.18, 3.32, 0.28, [{"text": d_sub, "font": MED, "size": 12.5, "color": MUTED}])
        rect(s, dx + 0.22, 3.52, 3.32, 0.012, HAIR)
        tb(s, dx + 0.22, 3.64, 3.32, 1.88, [{"text": d_body, "font": SANS, "size": 12.5, "color": BODY, "line": 1.30}])
        dx += 4.07

    footer_banner(s, "中国法人长期运行独立系统：先对齐数据结构、筑牢阿里云 PIPL 合规数仓与 AI 自研能力，即可与全球共享 AI 引擎。")
    folio(s, 9, total_pages)
    notes(s, "完全对应参考文稿《중국 전산팀이 먼저 채우고 있는 빈칸》：1. IT改组为DAX；2. 阿里云数仓+结合AI的内部BI+PIPL权限与日志合规；3. AI Coding Agent减少外包依赖；4. AI Agent帮助其他部门（有对策的越界 / AI Change Agent）；5. 对齐数据结构以共享AI引擎。")

    # =========================================================================
    # 10 收束：设计图与指南针已交出，砌砖立柱是我们的本分
    # =========================================================================
    s = new(prs)
    bg(s, DARK)
    rect(s, 0, 0, 0.09, H, RED)
    tb(s, 0.88, 0.62, 10, 0.3, [{"text": "结语  ·  100兆是今天一针一线填进去的冷静结果", "font": MED, "size": 13, "color": "C87A7A"}])

    summary_rows = [
        ("生存必要条件", "破除 3% 增速下的“企业麻风病”  ·  用 35% 返还筑起 AI 时代普通人养老的社会安全网"),
        ("一贯工程决胜", "一能一品（500ml→1000ml） · 先胜求战倒推工序成本 · 复制“十分之一价格 + 压倒性销量”"),
        ("3A 自动运转", "APP（PPO 与 AI 数字人） + Acare（免费健康锁客） + AZA（高频流通与威海物流枢纽）"),
        ("制度与现场姿态", "三无与上限线守住底层  ·  抛弃“野鸡埋雪”本位主义，践行 Chutzpah 精神与婴儿哲学"),
        ("中国 DAX 填空", "阿里云数仓（PIPL 合规） + AI Coding Agent 研发内化 + 跨部门 AI Agent 赋能与数据结构对齐"),
    ]
    sy = 1.15
    for s_left, s_right in summary_rows:
        rect(s, 0.88, sy, 11.55, 0.62, DARK_CARD)
        tb(s, 1.10, sy, 2.3, 0.62, [{"text": s_left, "font": MED, "size": 16, "color": RED}], valign="middle")
        tb(s, 3.55, sy, 8.7, 0.62, [{"text": s_right, "font": SANS, "size": 14, "color": "E0E0E0"}], valign="middle")
        sy += 0.72

    tb(
        s, 0.88, 4.98, 11.5, 0.68,
        [{"text": "会长交出了设计图与指南针，在图纸上砌砖立柱是现场的我们。", "font": MED, "size": 28, "color": WHITE}],
    )
    tb(
        s, 0.88, 5.76, 11.5, 0.82,
        [{"text": "100兆不是某一天像魔法一样自动打开的未来，而是今天我们拆改工序、打破惯例，\n用数据与 AI 一针一线填进计算表的冷静结果值。", "font": SANS, "size": 15.5, "color": "B5B5B5", "line": 1.35}],
    )
    tb(s, 0.88, 6.95, 9, 0.3, [{"text": "艾多美中国法人  ·  DAX（Digital & AI Transformation）", "font": SANS, "size": 12, "color": DIM}])
    folio(s, 10, total_pages, dark=True)
    notes(s, "收束全篇五章与中国DAX三大填空行动：会长已经交出设计图和指南针，在上面砌砖立柱是现场奔跑的我们的本分。")

    return prs


if __name__ == "__main__":
    out_path = Path(__file__).resolve().parent / "艾多美何以成为100兆韩元企业.pptx"
    build().save(str(out_path))
    print(out_path)
