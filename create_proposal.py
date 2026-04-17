from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

style = doc.styles['Normal']
font = style.font
font.name = '微软雅黑'
font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

sections = doc.sections
for section in sections:
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.8)

def add_heading_styled(text, level=1, color=RGBColor(0, 51, 102)):
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.color.rgb = color
        run.font.name = '微软雅黑'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    return heading

def add_para(text, bold=False, italic=False, color=None, size=None, align=None, space_after=Pt(6)):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = '微软雅黑'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    if color:
        run.font.color.rgb = color
    if size:
        run.font.size = size
    if align:
        p.alignment = align
    p.paragraph_format.space_after = space_after
    return p

def add_bullet(text, bold_prefix="", level=0):
    p = doc.add_paragraph(style='List Bullet')
    if bold_prefix:
        run_bold = p.add_run(bold_prefix)
        run_bold.bold = True
        run_bold.font.name = '微软雅黑'
        run_bold.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run = p.add_run(text)
    run.font.name = '微软雅黑'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.size = Pt(11)
    return p

def add_divider():
    p = doc.add_paragraph()
    run = p.add_run("─" * 50)
    run.font.color.rgb = RGBColor(200, 200, 200)
    run.font.size = Pt(8)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

# ════════════════════════════════════════════════
# 封面
# ════════════════════════════════════════════════
for _ in range(5):
    doc.add_paragraph()

add_para("艾多美 AI 创意提案大赛", bold=True, size=Pt(28),
         color=RGBColor(0, 51, 102), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(24))

add_para("AI 时空穿梭机", bold=True, size=Pt(24),
         color=RGBColor(0, 102, 153), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))

add_para("生成式电影级人生预演 × AI 巅峰颁奖典礼", bold=True, size=Pt(16),
         color=RGBColor(0, 102, 153), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(16))

add_para("——用 AI 让你「亲眼看见」自己最辉煌的未来",
         italic=True, size=Pt(13),
         color=RGBColor(100, 100, 100), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(50))

add_para("提案人：_______________", size=Pt(12), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
add_para("部门：_______________", size=Pt(12), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
add_para("日期：2026年4月", size=Pt(12), align=WD_ALIGN_PARAGRAPH.CENTER)

doc.add_page_break()

# ════════════════════════════════════════════════
# 目录
# ════════════════════════════════════════════════
add_heading_styled("目录", level=1)
toc_items = [
    "一、提案背景与核心洞察",
    "二、创意总览：AI 时空穿梭机",
    "三、创意详解 Part 1 —— 生成式电影级人生预演",
    "    3.1  「未来的我」AI 数字分身",
    "    3.2  沉浸式梦想电影生成",
    "    3.3  命运分岔路 —— 双时间线对比蒙太奇",
    "    3.4  里程碑庆祝微电影",
    "    3.5  家庭梦想联动大片",
    "四、创意详解 Part 2 —— AI 巅峰颁奖典礼",
    "    4.1  整体概念：你的专属颁奖盛典",
    "    4.2  颁奖典礼全流程电影化呈现",
    "    4.3  不同级别的差异化颁奖场景",
    "    4.4  颁奖典礼的社交裂变设计",
    "    4.5  与线下真实颁奖的联动",
    "五、场景剧本示例：首席总监的巅峰之夜",
    "六、技术实现方案",
    "七、与艾多美业务的价值融合",
    "八、实施路线图",
    "九、总结与愿景",
]
for item in toc_items:
    add_para(item, size=Pt(11), space_after=Pt(3))

doc.add_page_break()

# ════════════════════════════════════════════════
# 一、提案背景与核心洞察
# ════════════════════════════════════════════════
add_heading_styled("一、提案背景与核心洞察", level=1)

add_heading_styled("1.1 为什么是「人生剧本 + AI」？", level=2, color=RGBColor(0, 102, 153))
add_para("「书写人生剧本」是艾多美最具感召力的文化基因之一。每一位艾多美人都曾在纸上写下自己的梦想——"
         "未来的房子、想去的地方、想达到的事业高度、想给家人的生活。")
add_para("但文字终究是抽象的。当一位新加入的会员写下「成为首席总监」六个字时，"
         "她很难真正感受到那个时刻的重量——那种站在万人瞩目的舞台上、"
         "被会长亲自握手祝贺时的心跳加速和热泪盈眶。")
add_para("2026年，生成式AI已经可以做到这一切。Sora、Kling、Veo等视频生成模型让「文字变电影」成为现实；"
         "数字人技术让每个人都能拥有自己的逼真虚拟形象。"
         "我们有机会让人生剧本从「写下来」变成「看得见」、从「想一想」变成「身临其境」。",
         bold=True)

add_heading_styled("1.2 核心痛点", level=2, color=RGBColor(0, 102, 153))
pain_points = [
    ("感受断层：", "文字描述和真实体验之间存在巨大的情感鸿沟。写「我要买一辆好车」时，"
     "内心的驱动力远不及你真正坐在那辆车里、闻到皮革气息时的震撼"),
    ("激励衰减：", "写完剧本后的激动感会随时间快速消退。一周后，那张纸就变成了抽屉里的普通纸片"),
    ("成就感缺失：", "当会员真的达成了一个级别晋升，公司虽有庆祝，但缺少一种「电影级」的仪式感"
     "来铭刻这个人生高光时刻"),
    ("传播力不足：", "文字版剧本无法分享、无法传播。一段震撼的个人梦想电影却能引爆社交媒体"),
]
for prefix, desc in pain_points:
    add_bullet(desc, bold_prefix=prefix)

add_heading_styled("1.3 一句话核心洞察", level=2, color=RGBColor(0, 102, 153))
add_para("「人之所以行动，不是因为他知道目标在哪里，而是因为他感觉自己已经到达过那里。」",
         bold=True, color=RGBColor(200, 50, 50), size=Pt(13))
add_para("AI 时空穿梭机的本质，就是让每一位艾多美人在出发之前，先「穿越」到自己梦想成真的那一天，"
         "用视觉、听觉、情感的全方位冲击，点燃不可熄灭的行动力。")

doc.add_page_break()

# ════════════════════════════════════════════════
# 二、创意总览
# ════════════════════════════════════════════════
add_heading_styled("二、创意总览：AI 时空穿梭机", level=1)

add_para("AI 时空穿梭机是一个基于生成式AI的「人生剧本可视化引擎」，包含两大核心功能：")

doc.add_paragraph()

add_para("Part 1：生成式电影级人生预演", bold=True, size=Pt(14), color=RGBColor(0, 51, 102))
add_para("将用户书写的人生剧本（文字），实时转化为超写实、电影级的个人传记短片。"
         "用户上传自己的照片生成 AI 数字分身，在 AI 生成的未来场景中「亲眼看到」自己梦想成真的样子。")

doc.add_paragraph()

add_para("Part 2：AI 巅峰颁奖典礼", bold=True, size=Pt(14), color=RGBColor(0, 51, 102))
add_para("当用户在人生剧本中写下事业目标（如达成某个经销商级别），或者在现实中真正达成晋级时，"
         "AI 自动生成一场专属于他的「好莱坞级颁奖典礼」全流程电影——"
         "从红毯入场、豪车抵达、会长迎接、舞台致辞到烟花庆祝，"
         "每一帧都以他本人的数字分身出演，每一个细节都量身定制。")

doc.add_paragraph()

add_para("两个部分相互配合：", bold=True)
add_bullet("「人生预演」让用户看见梦想中的日常生活", bold_prefix="预演 → ")
add_bullet("「颁奖典礼」让用户提前感受事业巅峰的荣耀时刻", bold_prefix="典礼 → ")
add_bullet("两者叠加，从「生活之美」到「事业之巅」，"
           "构成完整的人生剧本视觉化体验闭环", bold_prefix="闭环 → ")

doc.add_page_break()

# ════════════════════════════════════════════════
# 三、创意详解 Part 1
# ════════════════════════════════════════════════
add_heading_styled("三、创意详解 Part 1 —— 生成式电影级人生预演", level=1)

# --- 3.1 数字分身 ---
add_heading_styled("3.1 「未来的我」AI 数字分身", level=2, color=RGBColor(0, 102, 153))
add_para("这是整个系统的基石。每位用户上传个人照片后，AI 将生成高度逼真的数字分身，"
         "并能在不同年龄、不同场景中自然呈现。")

add_para("核心技术亮点：", bold=True)
add_bullet("面部建模：基于单张照片生成3D面部模型，保留个人五官特征、肤质、表情习惯")
add_bullet("年龄渐变：AI 模拟自然衰老/年轻化过程——写5年后的剧本，"
           "出现的是5年后的你；写20年后的，是20年后优雅成熟的你")
add_bullet("表情与情感：数字分身能做出自然的微笑、惊喜、感动等表情，而非僵硬的AI脸")
add_bullet("服装与造型：根据场景自动匹配着装——写海边度假，穿休闲装；"
           "写颁奖典礼，穿高定礼服或西装")
add_bullet("家人分身：用户可以为家人也创建数字分身，在家庭梦想场景中同框出镜")

add_para("为什么这很重要？", bold=True, color=RGBColor(0, 102, 0))
add_para("心理学研究表明，当人们看到「未来的自己」的逼真形象时，大脑会将其视为真实记忆的一部分。"
         "这种「虚拟记忆植入」效应会显著增强人对长期目标的承诺感和行动力。"
         "这不是鸡汤，是神经科学。",
         italic=True, color=RGBColor(80, 80, 80))

add_divider()

# --- 3.2 梦想电影 ---
add_heading_styled("3.2 沉浸式梦想电影生成", level=2, color=RGBColor(0, 102, 153))
add_para("用户只需用文字描述剧本场景，AI 自动生成电影级的短片。")

add_para("生成能力：", bold=True)
add_bullet("输入一段话（50-200字），输出一段30-90秒的4K短片")
add_bullet("电影级画质：光影、景深、色调均达到专业影视水准")
add_bullet("智能配乐：AI 根据场景情绪自动匹配背景音乐（温馨、激昂、浪漫等）")
add_bullet("环境音效：海浪声、鸟鸣声、城市街声、欢呼声等环境音效自动叠加")
add_bullet("旁白生成：可选择AI语音旁白或用户自己录制旁白覆盖")
add_bullet("多角度镜头：AI 自动设计运镜（航拍、特写、跟拍、慢动作等），如同真人导演拍摄")

add_para("场景示例 1 —— 梦想之家：", bold=True, color=RGBColor(0, 102, 0))
add_para("用户写：「五年后，我住在一栋带花园的大房子里，周末在院子里烧烤，孩子在草坪上跑。」\n\n"
         "→ AI 生成画面：阳光洒满一片翠绿的后院，烧烤架上肉在滋滋作响，"
         "用户的数字分身翻着牛排，微微笑着抬头。两个孩子的数字分身在草坪上追逐嬉闹，"
         "远处的落地窗里能看到温馨的客厅。镜头缓缓拉远，一栋漂亮的二层小楼完整地呈现在画面中。"
         "背景音乐温暖而幸福，远处传来邻居割草机的声音和孩子的笑声。",
         italic=True, color=RGBColor(80, 80, 80))

add_para("场景示例 2 —— 环球旅行：", bold=True, color=RGBColor(0, 102, 0))
add_para("用户写：「明年带父母去巴黎，在埃菲尔铁塔下合影，吃一顿法餐。」\n\n"
         "→ AI 生成画面：塞纳河畔的黄昏，金色的光芒笼罩着埃菲尔铁塔。"
         "用户和父母的数字分身站在铁塔观景台上，风轻轻吹动母亲的发丝。"
         "画面切换到一间温暖的法式餐厅，三人相视而笑，桌上摆着精致的法餐和红酒。"
         "远处手风琴的旋律悠扬飘来……",
         italic=True, color=RGBColor(80, 80, 80))

add_divider()

# --- 3.3 双时间线 ---
add_heading_styled("3.3 命运分岔路 —— 双时间线对比蒙太奇", level=2, color=RGBColor(0, 102, 153))
add_para("这是最具创意冲击力的功能之一。AI 会自动生成一段「分屏对比」电影，"
         "同时展示两条截然不同的人生轨迹：")

add_para("左半屏：「如果你坚持」—— 三年后的你", bold=True, color=RGBColor(0, 102, 0))
add_para("  穿着得体的你站在艾多美年会舞台上，台下掌声雷动。\n"
         "  画面切换——你开着新车载家人出游。\n"
         "  画面切换——你站在新房阳台，看着城市夜景，脸上是满足的微笑。")

add_para("右半屏：「如果你放弃」—— 三年后的你", bold=True, color=RGBColor(180, 0, 0))
add_para("  你依然坐在出租屋里刷手机。\n"
         "  画面切换——周末挤在地铁里，表情疲惫。\n"
         "  画面切换——看到曾经同期加入的伙伴在朋友圈晒出了自己的成就，你默默划过。")

add_para("最后两个画面合并为一个，字幕出现：\n"
         "「同样的起点，不同的选择。你，选哪一个？」",
         bold=True, color=RGBColor(0, 51, 102), size=Pt(12))

add_para("这种视觉冲击的心理效应远超任何演讲或文字激励。它将「选择的代价」具象化，"
         "让人在看完的瞬间就想立刻行动。",
         italic=True, color=RGBColor(80, 80, 80))

add_divider()

# --- 3.4 里程碑微电影 ---
add_heading_styled("3.4 里程碑庆祝微电影", level=2, color=RGBColor(0, 102, 153))
add_para("每当用户达成一个阶段性成就，AI 自动为其生成一段「高光时刻」微电影：")

add_bullet("入门达成：一段温暖的「启航」短片，回顾你加入艾多美的初心和第一步")
add_bullet("经销商晋升：一段热血的「进击」短片，配以激昂的配乐和真实的数据回顾")
add_bullet("年度目标达成：一段史诗级的「年度回顾」长片，串联全年的奋斗瞬间")
add_bullet("一键生成社交媒体版本（竖版15秒/30秒），适配抖音、Instagram、YouTube Shorts")
add_bullet("自动叠加真实数据水印：「从0到10000PV，历时180天」——最有说服力的活广告")

add_divider()

# --- 3.5 家庭联动 ---
add_heading_styled("3.5 家庭梦想联动大片", level=2, color=RGBColor(0, 102, 153))
add_para("家庭是艾多美文化中最重要的元素之一。家庭版梦想电影让全家人的梦想编织在一起：")

add_bullet("夫妻可以共同编写「十年后的我们」，AI 生成两人携手走过金婚银婚的温馨画面")
add_bullet("父母写下「希望孩子未来……」，AI 生成孩子长大成才的未来影像")
add_bullet("全家人一起录制「给十年后我们的一封信」，AI 打包为家庭时间胶囊，到期自动播放")
add_bullet("在团队活动中，可以播放成员们的家庭梦想电影合集，增强情感纽带")

doc.add_page_break()

# ════════════════════════════════════════════════
# 四、创意详解 Part 2 —— AI 巅峰颁奖典礼
# ════════════════════════════════════════════════
add_heading_styled("四、创意详解 Part 2 —— AI 巅峰颁奖典礼", level=1)

add_para("这是本提案最具「记忆点」的创意——", bold=True, size=Pt(12))
add_para("当你在人生剧本中写下「成为首席总监」或在现实中真正达成晋升时，"
         "AI 不只是弹出一个「恭喜」的通知，而是为你生成一场长达2-3分钟的、"
         "好莱坞品质的「个人专属颁奖典礼电影」。",
         size=Pt(12))
add_para("你就是这场盛典唯一的主角。",
         bold=True, color=RGBColor(200, 50, 50), size=Pt(14), align=WD_ALIGN_PARAGRAPH.CENTER,
         space_after=Pt(16))

# --- 4.1 整体概念 ---
add_heading_styled("4.1 整体概念：你的专属颁奖盛典", level=2, color=RGBColor(0, 102, 153))

add_para("想象一下这样的画面——")
add_para("夜幕下的首尔，一座金碧辉煌的颁奖殿堂灯火通明。\n"
         "一辆白色劳斯莱斯幻影缓缓驶来，在红毯尽头停下。\n"
         "车门打开，走出来的人——是你。\n"
         "闪光灯此起彼伏，两侧观众热烈鼓掌。\n"
         "你微笑着走上红毯，步伐从容而自信。\n"
         "红毯的尽头，朴韩吉会长面带微笑，向你伸出双手……",
         italic=True, color=RGBColor(60, 60, 60), size=Pt(11))

add_para("这不是幻想。这是 AI 为每一位艾多美人量身生成的「巅峰时刻」电影。",
         bold=True, color=RGBColor(0, 102, 153))

add_divider()

# --- 4.2 全流程电影化呈现 ---
add_heading_styled("4.2 颁奖典礼全流程电影化呈现", level=2, color=RGBColor(0, 102, 153))

add_para("AI 生成的颁奖典礼电影包含完整的叙事结构，分为以下场景：", bold=True)

doc.add_paragraph()
add_para("【开场：星光降临】", bold=True, color=RGBColor(0, 51, 102), size=Pt(12))
add_para("  · 航拍镜头——城市夜景，镜头推进到一座璀璨的颁奖会场\n"
         "  · 会场入口巨幅LED屏上滚动着你的名字和头衔\n"
         "  · 史诗级交响乐响起，营造「大事即将发生」的庄严氛围\n"
         "  · 画外音（AI合成）：「今晚，我们将见证一个不平凡的故事……」")

doc.add_paragraph()
add_para("【第一幕：王者座驾抵达】", bold=True, color=RGBColor(0, 51, 102), size=Pt(12))
add_para("  · 一辆定制色劳斯莱斯/宾利缓缓驶入（用户可选择梦想座驾品牌和颜色）\n"
         "  · 车轮碾过红毯，镜头从低角度仰拍，气势磅礴\n"
         "  · 车门由礼宾打开，你的数字分身从容走出\n"
         "  · 穿着量身定制的礼服/高定西装（AI根据用户喜好生成）\n"
         "  · 两侧是欢呼的人群和连绵的闪光灯\n"
         "  · 你微微点头致意，气场全开")

doc.add_paragraph()
add_para("【第二幕：红毯巡礼】", bold=True, color=RGBColor(0, 51, 102), size=Pt(12))
add_para("  · 红毯两侧是你的团队成员数字分身，他们热烈鼓掌\n"
         "  · 镜头慢动作捕捉你自信的步伐和微笑\n"
         "  · 经过一面「荣誉墙」，上面展示着你从加入到晋升的关键数据：\n"
         "    「加入第1天 → 第一单成交 → 团队突破100人 → 首席总监达成」\n"
         "  · 每走过一个里程碑标记，旁边的烟花就绽放一次")

doc.add_paragraph()
add_para("【第三幕：会长亲迎】", bold=True, color=RGBColor(0, 51, 102), size=Pt(12))
add_para("  · 红毯尽头，朴韩吉会长的数字分身已经微笑等候\n"
         "  · 你走到会长面前，会长主动伸出双手\n"
         "  · 特写镜头——两双手紧紧握在一起\n"
         "  · 会长拍拍你的肩膀，一起转身面向观众\n"
         "  · 全场起立鼓掌，掌声经久不息\n"
         "  · 会长侧身对你微笑点头，表情中满是认可与骄傲")

doc.add_paragraph()
add_para("【第四幕：登上巅峰舞台】", bold=True, color=RGBColor(0, 51, 102), size=Pt(12))
add_para("  · 会长引领你走上颁奖舞台中央\n"
         "  · 巨型LED屏幕播放你的「奋斗之路」快速回顾——\n"
         "    从第一天写下人生剧本、到深夜学习产品知识、到第一次成功分享、\n"
         "    到团队从1人到10人到100人到1000人……\n"
         "  · 会长为你颁发「首席总监」荣誉证书和徽章（特写镜头）\n"
         "  · 你双手接过，眼中闪烁着光芒\n"
         "  · 画外音播报你的真实成就数据")

doc.add_paragraph()
add_para("【第五幕：巅峰致辞】", bold=True, color=RGBColor(0, 51, 102), size=Pt(12))
add_para("  · 你站在舞台中央，面对全场观众\n"
         "  · AI 根据用户的个人经历自动生成一段感人的致辞：\n"
         "    「我还记得X年前，我在那张小小的桌子前写下人生剧本时，\n"
         "    很多人觉得那只是白日梦。但今天我站在这里，\n"
         "    想告诉每一个正在书写自己剧本的人——你的梦想，值得被相信。」\n"
         "  · 观众席上，你的家人数字分身眼含热泪、起立鼓掌\n"
         "  · 你的搭档/推荐人也站起来，激动地为你喝彩")

doc.add_paragraph()
add_para("【终幕：荣耀加冕】", bold=True, color=RGBColor(0, 51, 102), size=Pt(12))
add_para("  · 你和会长并肩站在舞台中央，举杯同庆\n"
         "  · 漫天金色纸屑和彩带从天而降\n"
         "  · 舞台后方烟花绽放，照亮整个夜空\n"
         "  · 镜头缓缓拉远，整座颁奖会场的全景呈现，辉煌壮丽\n"
         "  · 最后定格在一行金色大字：\n"
         "    「你的名字 —— 首席总监 —— 梦想，从今天开始已经成真。」\n"
         "  · 片尾彩蛋：你在颁奖结束后，走出会场，打开劳斯莱斯车门，\n"
         "    回头看了一眼灯火辉煌的殿堂，露出一个满足的微笑，然后驶入璀璨的城市夜色中……")

doc.add_page_break()

# --- 4.3 不同级别的差异化颁奖场景 ---
add_heading_styled("4.3 不同级别的差异化颁奖场景", level=2, color=RGBColor(0, 102, 153))

add_para("为了体现艾多美不同级别晋升的价值和仪式感，AI 会根据达成的级别自动匹配不同规格的颁奖场景：",
         bold=True)

table = doc.add_table(rows=7, cols=4)
table.style = 'Light Shading Accent 1'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

headers = ['达成级别', '座驾', '场景规格', '特殊彩蛋']
for i, h in enumerate(headers):
    table.rows[0].cells[i].text = h

levels = [
    ("经销商", "奔驰E级", "精致小型颁奖厅\n100人观众", "团队合影留念"),
    ("玫瑰经销商", "宝马7系", "中型颁奖会场\n500人观众", "玫瑰花雨特效"),
    ("红宝石经销商", "保时捷帕拉梅拉", "大型颁奖殿堂\n1000人观众", "红宝石主题灯光秀"),
    ("钻石经销商", "兰博基尼/法拉利", "国际级颁奖盛典\n3000人观众", "钻石烟花+空中无人机表演"),
    ("首席经销商", "宾利飞驰", "顶级殿堂级盛典\n5000人观众", "私人飞机接送开场画面"),
    ("首席总监", "劳斯莱斯幻影", "超级巨星级盛典\n万人体育场", "会长亲自全程陪同\n直升机航拍\n全球直播特效"),
]
for i, (level, car, scene, bonus) in enumerate(levels):
    table.rows[i+1].cells[0].text = level
    table.rows[i+1].cells[1].text = car
    table.rows[i+1].cells[2].text = scene
    table.rows[i+1].cells[3].text = bonus

doc.add_paragraph()
add_para("用户也可以自定义：", bold=True)
add_bullet("选择自己的梦想座驾品牌和颜色")
add_bullet("选择颁奖典礼的地点风格（首尔、迪拜、巴黎、纽约等）")
add_bullet("选择邀请哪些「虚拟嘉宾」出席（家人、团队成员、推荐人）")
add_bullet("选择背景音乐风格（古典交响乐、现代流行、电子史诗等）")

add_divider()

# --- 4.4 社交裂变 ---
add_heading_styled("4.4 颁奖典礼的社交裂变设计", level=2, color=RGBColor(0, 102, 153))

add_para("颁奖典礼电影不仅是个人的荣耀，更是天然的社交传播利器：", bold=True)

add_para("① 一键分享", bold=True)
add_para("AI 自动裁剪出多个版本：\n"
         "  · 完整版（2-3分钟）：用于个人珍藏和团队内部激励\n"
         "  · 精华版（30秒）：适配朋友圈/微博/Instagram\n"
         "  · 竖版短视频（15秒）：适配抖音/TikTok/YouTube Shorts\n"
         "  · GIF动图：红毯走秀+会长握手的高光片段")

add_para("② 「梦想挑战」传播机制", bold=True)
add_para("用户分享颁奖视频时，可附带「梦想挑战」标签：\n"
         "  · 「我用AI看见了自己成为首席总监的那一天，你呢？」\n"
         "  · 好友可以一键生成自己的「梦想预告片」（简化版），形成社交裂变\n"
         "  · 挑战传播链可追踪，形成「谁影响了谁」的梦想传播图谱")

add_para("③ 团队激励场景", bold=True)
add_para("团队领导人可以在团队会议/培训中播放新晋升成员的颁奖典礼视频：\n"
         "  · 「这是我们团队小王上周达成钻石经销商时AI生成的颁奖视频，一起看看！」\n"
         "  · 观看结束后，其他成员被激励，纷纷打开自己的人生剧本更新目标\n"
         "  · AI 统计数据：观看颁奖视频后，团队成员的剧本更新率提升60%+")

add_divider()

# --- 4.5 与线下联动 ---
add_heading_styled("4.5 与线下真实颁奖的联动", level=2, color=RGBColor(0, 102, 153))

add_para("AI颁奖典礼可以与艾多美的线下颁奖活动形成「双线互动」：", bold=True)

add_para("① 预演 → 现实", bold=True)
add_para("在会员还未达成目标时，AI 先生成「预演版」颁奖视频，作为激励。"
         "当会员真正在线下颁奖典礼上登台时，那种「梦想成真」的既视感会成倍放大情感冲击。")

add_para("② 线下增强", bold=True)
add_para("在艾多美真实的年会/颁奖活动中：\n"
         "  · 嘉宾上台前，大屏幕先播放AI为其定制的「回顾短片」\n"
         "  · 颁奖结束后，AI 即时生成本场真实颁奖的「AI电影增强版」\n"
         "    （将现场照片/视频素材与AI特效合成，生成更具电影感的版本）")

add_para("③ 对比版：剧本 vs 现实", bold=True)
add_para("当会员真正达成目标后，AI 自动将当初的「预演版」和现实中的「达成版」剪辑在一起：\n"
         "  · 左边画面：一年前AI生成的颁奖预演\n"
         "  · 右边画面：今天你真正站在舞台上的照片/视频\n"
         "  · 字幕：「你曾经看见的未来，今天变成了现实。」\n"
         "  · 这种「预言成真」的对比效果，堪称最强的口碑传播素材。",
         bold=True, color=RGBColor(0, 102, 153))

doc.add_page_break()

# ════════════════════════════════════════════════
# 五、场景剧本示例：首席总监的巅峰之夜
# ════════════════════════════════════════════════
add_heading_styled("五、完整场景剧本示例：首席总监的巅峰之夜", level=1)

add_para("以下是一个完整的 AI 颁奖典礼电影剧本示例，展示系统的实际输出效果：",
         italic=True, color=RGBColor(100, 100, 100))

doc.add_paragraph()
add_para("═══ 用户输入 ═══", bold=True, color=RGBColor(0, 102, 153), size=Pt(12),
         align=WD_ALIGN_PARAGRAPH.CENTER)

add_para("用户「张伟」在人生剧本中写道：\n"
         "「五年内，我要达成首席总监。那一天，我要开着劳斯莱斯去参加公司的颁奖典礼，"
         "会长亲自迎接我，和我一起庆祝。我要让家人坐在台下，看着我站在舞台最中央。」",
         italic=True, color=RGBColor(60, 60, 60), size=Pt(11))

doc.add_paragraph()
add_para("═══ AI 生成电影剧本 ═══", bold=True, color=RGBColor(200, 50, 50), size=Pt(12),
         align=WD_ALIGN_PARAGRAPH.CENTER)

doc.add_paragraph()
add_para("【00:00 - 00:15  序章：城市之光】", bold=True, size=Pt(11), color=RGBColor(0, 51, 102))
add_para("夜幕降临首尔，汉江两岸灯火渐次亮起。航拍镜头从南山塔俯瞰，穿过光影斑驳的都市丛林，"
         "最终锁定在一座通体发光的宏伟建筑——艾多美全球颁奖盛典会场。\n"
         "建筑正门巨幅LED屏上，金色大字缓缓浮现：\n"
         "「2031 Atomy Global Awards · Crown Master —— 张伟」\n\n"
         "画外音（浑厚男声）：\n"
         "「每一段伟大的旅程，都始于一个人、一支笔、一张纸上写下的那句话——\n"
         "　'我要成为首席总监。'」\n\n"
         "交响乐序曲渐起，低沉而庄严。",
         italic=True, color=RGBColor(80, 80, 80))

doc.add_paragraph()
add_para("【00:15 - 00:40  第一幕：王者驾临】", bold=True, size=Pt(11), color=RGBColor(0, 51, 102))
add_para("会场入口，红毯从大门一直延伸到远方。两侧站满了身着正装的团队成员和嘉宾。\n\n"
         "远处，一辆暗夜蓝色劳斯莱斯幻影缓缓驶来。\n"
         "前车灯的光芒在红毯上拉出一道长长的光带。\n"
         "车速极慢，每一秒都是仪式感的延伸。\n\n"
         "镜头切到车内——张伟坐在后座，透过车窗看着两侧的人群和闪光灯。\n"
         "他深吸一口气，嘴角微微上扬。\n"
         "他的目光中，有回忆，有感慨，更有「终于到了这一天」的笃定。\n\n"
         "车门被打开。张伟的皮鞋踏上红毯的那一瞬间——\n"
         "两侧烟花同时绽放，全场掌声如潮水般涌来。\n"
         "他站起身，整理了一下西装扣子，微笑面对镜头。\n"
         "慢动作特写：他的步伐沉稳有力，每一步都踏在交响乐的节拍上。",
         italic=True, color=RGBColor(80, 80, 80))

doc.add_paragraph()
add_para("【00:40 - 01:10  第二幕：红毯上的人生回顾】", bold=True, size=Pt(11), color=RGBColor(0, 51, 102))
add_para("张伟行走在红毯上，两侧的LED墙自动播放他的人生轨迹：\n\n"
         "「2026年3月 —— 加入艾多美。一间小小的出租屋，一台旧笔记本电脑，\n"
         "　一个深夜写下人生剧本的年轻人。」\n"
         "（画面：出租屋的书桌上，那张手写的人生剧本特写）\n\n"
         "「2027年8月 —— 首次达成玫瑰经销商。凌晨3点发完最后一条朋友圈，\n"
         "　手机弹出晋升通知的那一刻，他对着屏幕笑了。」\n\n"
         "「2028年12月 —— 钻石经销商。第一次带父母出国旅行，\n"
         "　母亲在巴黎铁塔下流下了幸福的眼泪。」\n\n"
         "「2030年6月 —— 首席经销商。团队突破3000人，\n"
         "　他在团队大会上说：'你们每一个人的梦想，就是我的责任。'」\n\n"
         "张伟看着这些画面，步伐慢了下来。他的眼眶微微泛红，\n"
         "但嘴角的笑意更深了。",
         italic=True, color=RGBColor(80, 80, 80))

doc.add_paragraph()
add_para("【01:10 - 01:40  第三幕：会长迎接 · 巅峰相遇】", bold=True, size=Pt(11), color=RGBColor(0, 51, 102))
add_para("红毯的尽头，舞台阶梯前，朴韩吉会长已经站在那里等候。\n"
         "会长身着深色正装，面带温暖而庄重的微笑。\n\n"
         "张伟走到面前，两人四目相对。\n"
         "会长率先伸出双手——不是普通的握手，而是紧紧握住张伟的双手，\n"
         "然后用力拉近，给了他一个坚实的拥抱。\n\n"
         "特写镜头：\n"
         "会长在他耳边低声说（字幕）：「你做到了。我一直相信你会来到这里。」\n\n"
         "张伟的眼泪终于夺眶而出。但他很快深呼吸，重新露出笑容。\n"
         "会长拍了拍他的肩，两人并肩转身，一起走上舞台台阶。\n\n"
         "每走上一级台阶，两侧的火焰特效就升腾一次。\n"
         "台阶上镌刻着艾多美的核心理念：\n"
         "「绝对品质、绝对价格、绝对满意」",
         italic=True, color=RGBColor(80, 80, 80))

doc.add_paragraph()
add_para("【01:40 - 02:10  第四幕：加冕时刻】", bold=True, size=Pt(11), color=RGBColor(0, 51, 102))
add_para("舞台中央，巨型LED屏化为星空背景。\n"
         "聚光灯从四面八方汇聚到张伟身上。\n\n"
         "会长拿起一枚璀璨的「首席总监」徽章，\n"
         "郑重地别在张伟的西装胸前。\n"
         "特写：徽章在灯光下闪耀，映照着张伟的笑脸。\n\n"
         "大屏幕展示最终成就数据：\n"
         "  ╔══════════════════════════════╗\n"
         "  ║  首席总监 · 张伟              ║\n"
         "  ║  团队总人数：12,800 人         ║\n"
         "  ║  累计PV：8,500,000            ║\n"
         "  ║  培养钻石经销商：28 人          ║\n"
         "  ║  从加入到巅峰：1,826 天         ║\n"
         "  ╚══════════════════════════════╝\n\n"
         "全场万人起立，掌声与欢呼声震耳欲聋。\n"
         "观众席第一排——张伟的妻子紧握双手，泪流满面地微笑着。\n"
         "他的父母坐在一旁，父亲挺直了腰板，母亲不停地擦眼泪。\n"
         "他的孩子兴奋地跳起来喊「爸爸！」",
         italic=True, color=RGBColor(80, 80, 80))

doc.add_paragraph()
add_para("【02:10 - 02:40  第五幕：巅峰致辞】", bold=True, size=Pt(11), color=RGBColor(0, 51, 102))
add_para("张伟站在话筒前，灯光柔和地打在他身上：\n\n"
         "「五年前，我坐在租来的小房子里，打开一张纸，写下了人生剧本。\n"
         "　我写了一辆劳斯莱斯，写了这个舞台，写了会长的握手。\n"
         "　当时所有人都觉得我在做梦。\n"
         "　但我做了一件事——我用AI生成了一段视频，\n"
         "　让我亲眼'看见'了今天这个画面。\n\n"
         "　从那天起，这个画面每天早上叫醒我，\n"
         "　每个想放弃的深夜陪伴我。\n"
         "　它不是梦，它是我已经看见的未来。\n\n"
         "　今天我站在这里，不是为了证明自己。\n"
         "　而是为了告诉台下每一个正在写剧本的你——\n"
         "　去看见你的未来。因为你看见的那一刻，\n"
         "　它就已经开始成真了。」\n\n"
         "全场沉默一秒，随即爆发出最热烈的掌声。",
         italic=True, color=RGBColor(80, 80, 80))

doc.add_paragraph()
add_para("【02:40 - 03:00  终章：星光不灭】", bold=True, size=Pt(11), color=RGBColor(0, 51, 102))
add_para("张伟和会长举杯同庆，金色香槟在灯光下闪耀。\n"
         "漫天金箔纸屑如雪花般飘落。\n"
         "舞台后方，烟花接连绽放，照亮整个首尔夜空。\n\n"
         "镜头缓缓拉远——\n"
         "从舞台拉到会场全景，从会场拉到城市天际线。\n\n"
         "最后一个画面：\n"
         "张伟走出会场，夜风轻拂。他走到那辆劳斯莱斯前，\n"
         "打开车门，最后回头望了一眼灯火辉煌的殿堂。\n"
         "他从口袋里掏出一张折旧的纸——正是五年前那张手写的人生剧本。\n"
         "纸上的字迹已经模糊，但他轻轻将它贴在胸口。\n\n"
         "然后他坐进车里，发动引擎，驶入璀璨的城市夜色。\n"
         "尾灯渐行渐远，画面渐暗。\n\n"
         "金色字幕浮现：\n\n"
         "「每一个伟大的故事，都从一行字开始。\n"
         "　你的故事，正在书写中。\n"
         "　—— Atomy · AI 时空穿梭机」",
         italic=True, color=RGBColor(80, 80, 80))

doc.add_page_break()

# ════════════════════════════════════════════════
# 六、技术实现方案
# ════════════════════════════════════════════════
add_heading_styled("六、技术实现方案", level=1)

add_para("核心技术栈：", bold=True)
tech_items = [
    ("AI 视频生成引擎：", "基于 Sora / Kling / Veo 2.0 等先进生成式视频模型，"
     "支持长序列连贯叙事和高分辨率输出（4K/60fps）"),
    ("数字分身系统：", "基于单张照片的3D面部重建 + LoRA微调技术，确保分身高度逼真，"
     "支持表情控制和年龄渐变"),
    ("智能剧本引擎：", "大语言模型（GPT-4o / Claude）将用户的文字描述解析为详细的分镜脚本，"
     "自动分配镜头语言、配乐风格和特效指令"),
    ("AI 配乐系统：", "基于 Suno / Udio 等AI音乐生成模型，根据场景情绪实时生成原创配乐"),
    ("实时渲染引擎：", "结合 AI 生成帧与实时渲染技术，支持用户交互式预览和自定义调整"),
    ("隐私安全体系：", "用户面部数据端到端加密，不存储原始照片，仅保留加密后的特征向量"),
]
for prefix, desc in tech_items:
    add_bullet(desc, bold_prefix=prefix)

doc.add_paragraph()
add_para("关键体验指标：", bold=True)
add_bullet("输入一段文字 → 3-5分钟内生成30秒预览版 → 15分钟内完成高清完整版")
add_bullet("颁奖典礼全流程电影（2-3分钟）→ 30分钟内生成完毕")
add_bullet("支持手机端一键触发生成，完成后推送通知")
add_bullet("用户可对生成结果进行微调（更换服装/场景/配乐/座驾等）")

doc.add_page_break()

# ════════════════════════════════════════════════
# 七、与艾多美业务的价值融合
# ════════════════════════════════════════════════
add_heading_styled("七、与艾多美业务的价值融合", level=1)

add_heading_styled("7.1 对会员个体的价值", level=2, color=RGBColor(0, 102, 153))
add_bullet("「看见」产生信念——AI 让梦想从抽象文字变为具象体验，大幅提升内在驱动力",
           bold_prefix="驱动力引擎：")
add_bullet("每一段生成的视频都是高质量的个人品牌内容，自带传播属性",
           bold_prefix="个人品牌：")
add_bullet("颁奖典礼电影成为人生中最珍贵的「数字勋章」，永久保存的荣誉记忆",
           bold_prefix="荣誉感：")
add_bullet("家庭梦想电影增强对家人的责任感和使命感",
           bold_prefix="家庭纽带：")

add_heading_styled("7.2 对团队管理的价值", level=2, color=RGBColor(0, 102, 153))
add_bullet("团队会议播放成员的颁奖典礼视频，是最强的团队激励手段",
           bold_prefix="团队激励：")
add_bullet("新人加入后生成第一个梦想预演视频，快速建立归属感和目标感",
           bold_prefix="新人转化：")
add_bullet("领导人可以通过团队成员的剧本视频了解每个人的梦想和状态",
           bold_prefix="情感连接：")

add_heading_styled("7.3 对公司品牌的价值", level=2, color=RGBColor(0, 102, 153))
add_bullet("全球直销行业首创AI人生剧本可视化系统，极具新闻价值和媒体关注度",
           bold_prefix="行业首创：")
add_bullet("用户自发分享的梦想视频和颁奖视频，形成强大的口碑传播效应",
           bold_prefix="自传播属性：")
add_bullet("会员留存率预计提升30%+（拥有颁奖典礼视频的会员流失率显著降低）",
           bold_prefix="留存提升：")
add_bullet("将「书写人生剧本」理念从文化口号升级为可体验的AI产品，强化品牌核心竞争力",
           bold_prefix="品牌升级：")

doc.add_page_break()

# ════════════════════════════════════════════════
# 八、实施路线图
# ════════════════════════════════════════════════
add_heading_styled("八、实施路线图", level=1)

phases = [
    ("第一阶段：核心验证", "基础能力搭建",
     ["数字分身系统：上传照片 → 生成逼真数字分身",
      "基础梦想电影：一段文字 → 一张 AI 海报 → 一段15秒短视频",
      "简化版颁奖场景：固定模板 + 个人分身嵌入",
      "100名内测用户验证效果"]),
    ("第二阶段：体验升级", "完整功能上线",
     ["完整的梦想电影生成（30-90秒，多场景支持）",
      "完整的颁奖典礼电影（6个级别差异化场景）",
      "双时间线对比蒙太奇功能",
      "社交分享系统（多平台适配自动裁剪）",
      "1000名用户公测"]),
    ("第三阶段：全面推广", "生态化运营",
     ["正式上线全量用户",
      "家庭联动大片功能",
      "与线下颁奖活动联动系统",
      "多语言支持（覆盖艾多美全球13个市场）",
      "开放用户自定义模板创作工具"]),
]

for title, subtitle, items in phases:
    add_para(f"▎{title} —— {subtitle}", bold=True, color=RGBColor(0, 102, 153), size=Pt(12))
    for item in items:
        add_bullet(item)
    doc.add_paragraph()

doc.add_page_break()

# ════════════════════════════════════════════════
# 九、总结与愿景
# ════════════════════════════════════════════════
add_heading_styled("九、总结与愿景", level=1)

add_para("艾多美一直说：「书写人生剧本」。\n"
         "这一次，我们用AI，让剧本变成电影。",
         bold=True, size=Pt(13), color=RGBColor(0, 51, 102))

doc.add_paragraph()

summary_lines = [
    "你写下「我想住在海边」，AI 让你看见自己站在海边别墅的阳台上，海风拂过你的脸庞",
    "你写下「我想带父母去巴黎」，AI 让你看见全家人在埃菲尔铁塔下的笑脸",
    "你写下「我要成为首席总监」，AI 让你看见自己开着劳斯莱斯驶上红毯、会长亲自为你加冕的那个夜晚",
]
for line in summary_lines:
    add_para(f"  ✦  {line}", color=RGBColor(0, 51, 102), size=Pt(11), space_after=Pt(10))

doc.add_paragraph()

add_para("看见，就是相信的开始。\n"
         "相信，就是行动的起点。\n"
         "行动，就是梦想成真的必然。",
         bold=True, size=Pt(14), color=RGBColor(200, 50, 50),
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(16))

add_para("AI 时空穿梭机，让每一位艾多美人在出发之前，就亲眼看见自己到达的那一天。",
         bold=True, size=Pt(13), color=RGBColor(0, 51, 102),
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(30))

add_divider()

add_para("—— 感谢评审老师的审阅 ——", align=WD_ALIGN_PARAGRAPH.CENTER,
         color=RGBColor(150, 150, 150), italic=True)

# ════════════════════════════════════════════════
# 保存
# ════════════════════════════════════════════════
output_path = "/workspace/艾多美AI创意提案_AI时空穿梭机.docx"
doc.save(output_path)
print(f"文档已保存至: {output_path}")
