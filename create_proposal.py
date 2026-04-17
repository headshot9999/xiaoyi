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

# ============ 封面 ============
for _ in range(6):
    doc.add_paragraph()

add_para("艾多美 AI 创意提案大赛", bold=True, size=Pt(28),
         color=RGBColor(0, 51, 102), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(20))

add_para("AI × 人生剧本", bold=True, size=Pt(22),
         color=RGBColor(0, 102, 153), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(8))

add_para("——让每一位艾多美人的梦想「可看见、可运行、可实现」",
         italic=True, size=Pt(14),
         color=RGBColor(100, 100, 100), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(40))

add_para("提案人：_______________", size=Pt(12), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
add_para("部门：_______________", size=Pt(12), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
add_para("日期：2026年4月", size=Pt(12), align=WD_ALIGN_PARAGRAPH.CENTER)

doc.add_page_break()

# ============ 目录 ============
add_heading_styled("目录", level=1)
toc_items = [
    "一、提案背景与洞察",
    "二、核心理念：AI 人生剧本 3.0",
    "三、五大创意模块详解",
    "    模块1：时空穿梭机 —— 生成式电影级人生预演",
    "    模块2：人生自动驾驶 —— AI Agent 目标执行引擎",
    "    模块3：实时进度追踪 —— 梦想仪表盘",
    "    模块4：剧本社交宇宙 —— 梦想共振网络（新增）",
    "    模块5：AI 人生导演 —— 智能剧本优化器（新增）",
    "四、技术架构概览",
    "五、与艾多美业务的深度融合",
    "六、实施路线图",
    "七、预期效果与价值",
    "八、总结"
]
for item in toc_items:
    add_para(item, size=Pt(12), space_after=Pt(4))

doc.add_page_break()

# ============ 一、提案背景与洞察 ============
add_heading_styled("一、提案背景与洞察", level=1)

add_heading_styled("1.1 行业趋势", level=2, color=RGBColor(0, 102, 153))
add_para("在 AI 技术飞速发展的2026年，生成式AI已经从「文本对话」进化到「多模态创造」的新阶段。"
         "Sora、Kling等视频生成模型让普通人也能创造电影级内容；"
         "AI Agent 技术让AI从「回答问题」升级为「主动执行任务」。"
         "这为艾多美的「人生剧本」理念注入了前所未有的想象空间。")

add_heading_styled("1.2 痛点洞察", level=2, color=RGBColor(0, 102, 153))

pain_points = [
    ("看不见：", "传统的人生剧本停留在文字层面，缺乏沉浸感和感染力，会员难以产生强烈的情感共鸣"),
    ("跑不动：", "剧本写完后缺乏执行路径，目标与日常行动之间存在断层，「写了就忘」现象普遍"),
    ("没反馈：", "缺少实时追踪机制，会员不知道自己距离梦想还有多远，容易中途放弃"),
    ("太孤独：", "逐梦之路缺少同伴的激励和见证，个体的坚持力有限"),
]
for prefix, desc in pain_points:
    add_bullet(desc, bold_prefix=prefix)

add_heading_styled("1.3 核心洞察", level=2, color=RGBColor(0, 102, 153))
add_para("「人生剧本」不应该只是一张纸上的愿望清单，它应该是一部可以「预览」的电影、一套可以「运行」的程序、"
         "一个可以「追踪」的仪表盘、一张可以「共振」的社交网络。",
         bold=True, color=RGBColor(0, 102, 153), size=Pt(12))

doc.add_page_break()

# ============ 二、核心理念 ============
add_heading_styled("二、核心理念：AI 人生剧本 3.0", level=1)

add_para("我们提出「AI 人生剧本 3.0」概念，将艾多美的人生剧本从三个维度进行革命性升级：")

versions = [
    ("人生剧本 1.0（现状）", "手写文字 → 愿望清单 → 静态展示"),
    ("人生剧本 2.0（进化）", "AI生成 → 多模态呈现 → 沉浸体验"),
    ("人生剧本 3.0（革命）", "AI驱动 → 自动执行 → 社交共振 → 闭环实现"),
]

table = doc.add_table(rows=4, cols=2)
table.style = 'Light Shading Accent 1'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

hdr_cells = table.rows[0].cells
hdr_cells[0].text = '版本'
hdr_cells[1].text = '特征'
for i, (ver, feat) in enumerate(versions):
    row_cells = table.rows[i+1].cells
    row_cells[0].text = ver
    row_cells[1].text = feat

doc.add_paragraph()

add_para("一句话总结：", bold=True)
add_para("「用AI让梦想可看见、可运行、可追踪、可共振、可优化——从写剧本到活在剧本里。」",
         bold=True, italic=True, color=RGBColor(200, 50, 50), size=Pt(13))

doc.add_page_break()

# ============ 三、五大创意模块详解 ============
add_heading_styled("三、五大创意模块详解", level=1)

# --- 模块1 ---
add_heading_styled("模块1：🎬 时空穿梭机 —— 生成式电影级人生预演", level=2, color=RGBColor(0, 102, 153))

add_para("【原始创意】", bold=True, color=RGBColor(100, 100, 100))
add_para("利用生成式视频技术，将文字剧本转化为电影级个人传记短片。")

add_para("【升级创意】", bold=True, color=RGBColor(200, 50, 50))

add_para("① 「未来的我」数字分身", bold=True)
add_para("用户上传一张照片，AI 基于面部特征生成不同年龄阶段的数字分身。"
         "当你书写10年后的人生剧本时，画面中出现的是经过自然老化处理的「真实的未来你」，"
         "而不是一个虚假的AI人脸。这种真实感会产生强烈的情感冲击。")

add_para("② 五感沉浸式体验", bold=True)
add_para("不仅仅是视频——配合VR/AR设备，加入环境音效（海浪声、鸟鸣声）、"
         "甚至气味模拟（咖啡香、花香），让用户在「时空穿梭」中获得全方位的沉浸感。"
         "即使没有VR设备，手机端也可以通过裸眼3D效果和空间音频提供轻量级沉浸体验。")

add_para("③ 时间线对比蒙太奇", bold=True)
add_para("AI 自动生成两条时间线的对比视频：「如果你坚持」vs「如果你放弃」。"
         "左边画面是你站在梦想的舞台上接受表彰，右边是你依然在原地踏步。"
         "这种视觉对比产生的心理冲击远超任何鸡汤文字。")

add_para("④ 里程碑庆祝电影", bold=True)
add_para("每当用户达成一个阶段性目标（如晋升钻石经销商），AI 自动为其生成一段30秒的"
         "「人生高光时刻」短片，配以史诗级配乐和真实数据回顾。"
         "这段视频可以一键分享到社交媒体，成为最有感染力的「活广告」。")

add_para("⑤ 家庭剧本联动", bold=True)
add_para("家庭成员可以共同编写「家庭人生剧本」，AI 将生成包含所有家庭成员数字分身的"
         "家庭梦想电影。孩子的成长、父母的旅行、全家的新居——所有人的梦想编织在一起。")

add_para("【场景示例】", bold=True, color=RGBColor(0, 102, 0))
add_para("小美写下：「三年后，我带着爸妈去济州岛旅行，住在海景酒店，一起看日出。」\n"
         "→ AI 生成：一段60秒的4K短片，清晨的济州岛海岸线，阳光穿过云层洒在海面上，"
         "画面中是小美和父母的数字分身，三人站在阳台上，海风轻拂发丝，远处传来海浪声。"
         "妈妈转过头微笑，爸爸端着咖啡……\n"
         "→ 小美看到这一幕，眼眶湿润。这不再是一个梦，而是一个「已经看见的未来」。",
         italic=True, color=RGBColor(80, 80, 80))

doc.add_page_break()

# --- 模块2 ---
add_heading_styled("模块2：🤖 人生自动驾驶 —— AI Agent 目标执行引擎", level=2, color=RGBColor(0, 102, 153))

add_para("【原始创意】", bold=True, color=RGBColor(100, 100, 100))
add_para("引入AI Agent，将人生目标自动拆解为可执行的业务指标和行动步骤。")

add_para("【升级创意】", bold=True, color=RGBColor(200, 50, 50))

add_para("① 智能目标拆解树", bold=True)
add_para("用户输入一个大目标（如「三年内实现财务自由」），AI Agent 不是简单地列清单，"
         "而是构建一棵动态的「目标拆解树」：\n"
         "  · 第一层：财务自由 = 被动收入 ≥ 生活支出\n"
         "  · 第二层：被动收入来源分析（艾多美佣金、投资收益、副业收入）\n"
         "  · 第三层：艾多美佣金路径（当前等级→目标等级→需要的PV→团队规模）\n"
         "  · 第四层：每月/每周/每日具体行动（拜访客户数、培训次数、社交媒体发布数）\n"
         "每一层都有具体的数字和时间节点，并且会根据实际进度动态调整。")

add_para("② AI 智能教练系统", bold=True)
add_para("AI 不只是拆解目标，还会像教练一样主动推送行动建议：\n"
         "  · 早晨提醒：「今天的目标是联系3位潜在客户，我已经帮你筛选了5位高意向名单」\n"
         "  · 实时建议：「你今天分享的产品帖子互动率较低，建议改用视频+故事的形式，参考模板→」\n"
         "  · 晚间复盘：「今天完成了2/3的拜访目标，距离本月里程碑还差15个PV，加油！」")

add_para("③ 自动化社交媒体运营", bold=True)
add_para("AI Agent 可以根据用户的剧本目标，自动化完成以下工作（在用户授权和合规前提下）：\n"
         "  · 自动生成产品推荐文案（基于用户个人风格和受众分析）\n"
         "  · 智能排期发布（根据粉丝活跃时段优化发布时间）\n"
         "  · 自动回复评论和私信中的常见问题\n"
         "  · 生成每周社交媒体运营报告")

add_para("④ 「剧本日历」同步", bold=True)
add_para("AI 将拆解后的行动步骤自动同步到用户的日历（Google Calendar/Apple Calendar），"
         "每一个微小的行动都成为日程上的一个待办事项。人生剧本不再是墙上的装饰品，"
         "而是渗透到每天日程里的行动指南。")

add_para("⑤ 团队协同驾驶", bold=True)
add_para("上线可以看到下线的目标进度（在授权范围内），AI 会自动建议团队协作方案：\n"
         "  · 「你的团队成员小李本月PV差距较大，建议安排一次一对一辅导」\n"
         "  · 「团队中有3位成员对护肤品类感兴趣，建议组织一次产品知识分享会」\n"
         "  · 自动生成团队周报和月度复盘报告")

doc.add_page_break()

# --- 模块3 ---
add_heading_styled("模块3：📊 实时进度追踪 —— 梦想仪表盘", level=2, color=RGBColor(0, 102, 153))

add_para("【原始创意】", bold=True, color=RGBColor(100, 100, 100))
add_para("AI 接入用户的财务和事业数据，实时分析进度。")

add_para("【升级创意】", bold=True, color=RGBColor(200, 50, 50))

add_para("① 3D 可视化梦想星球", bold=True)
add_para("将传统的进度条升级为「梦想星球」——一个3D可交互的虚拟星球。\n"
         "星球上的每一座建筑代表一个人生目标：\n"
         "  · 「梦想之家」从地基开始，随着你的存款增长逐渐建造\n"
         "  · 「事业高塔」随着你的团队规模扩大而不断加盖楼层\n"
         "  · 「旅行地标」每去一个地方，星球上就多一个地标建筑\n"
         "  · 星球上有昼夜变化、天气系统，你的活跃度影响天气（高产的日子是晴天）")

add_para("② AI 数据分析师", bold=True)
add_para("AI 不只告诉你「完成了多少」，更会分析「为什么」和「怎么办」：\n"
         "  · 趋势预测：「按照当前增长速率，你将在14个月后达成钻石经销商目标，"
         "比原计划提前2个月」\n"
         "  · 瓶颈诊断：「你的新客户转化率（18%）低于优秀经销商平均水平（32%），"
         "建议优化首次接触话术，参考→」\n"
         "  · 机会发现：「本月健康食品类目增长37%，建议增加该类产品推广比重」")

add_para("③ 情感化进度反馈", bold=True)
add_para("进度追踪不只是冰冷的数字，AI 会根据进度阶段给出不同的情感化反馈：\n"
         "  · 起步期（0-20%）：AI 讲述类似背景会员的成功故事作为激励\n"
         "  · 爬坡期（20-60%）：AI 分析具体困难并提供针对性策略\n"
         "  · 突破期（60-90%）：AI 播放你之前录制的「给未来自己的一封信」\n"
         "  · 冲刺期（90-100%）：AI 预先生成庆祝视频，让你提前感受成功的喜悦")

add_para("④ 多维度生活平衡轮", bold=True)
add_para("不仅追踪事业目标，还关注人生的全面发展：\n"
         "  · 健康指数（运动、饮食、体检数据）\n"
         "  · 学习成长（培训参与、读书、技能提升）\n"
         "  · 家庭幸福（家庭活动、亲子时光）\n"
         "  · 社交关系（团队凝聚力、朋友互动）\n"
         "  · 财务状况（收入、储蓄、投资）\n"
         "以雷达图形式呈现，AI 会提醒你关注低分维度，避免「赚了钱却失了生活」。")

doc.add_page_break()

# --- 模块4（新增） ---
add_heading_styled("模块4：🌐 剧本社交宇宙 —— 梦想共振网络（全新创意）", level=2, color=RGBColor(0, 102, 153))

add_para("这是在原有创意基础上全新提出的模块，旨在解决「逐梦路上太孤独」的痛点。",
         italic=True, color=RGBColor(100, 100, 100))

add_para("① 梦想匹配算法", bold=True)
add_para("AI 分析所有会员的人生剧本，自动匹配「梦想相似」的伙伴：\n"
         "  · 「你们都想在2028年去欧洲旅行，要不要组队一起实现？」\n"
         "  · 「有12位会员和你一样想在今年考取营养师资格证，加入学习圈？」\n"
         "  · 志同道合的人组成「梦想共振小组」，互相激励和监督")

add_para("② 剧本直播间", bold=True)
add_para("会员可以在「剧本直播间」实时分享自己的人生剧本和追梦进度：\n"
         "  · AI 自动生成直播间的视觉背景（基于你的梦想场景）\n"
         "  · 其他会员可以为你的梦想「点亮星星」表示支持\n"
         "  · 达成里程碑时自动触发「全网庆祝」动效——所有关注你的人都会收到通知")

add_para("③ 时间胶囊社区", bold=True)
add_para("每位会员可以录制「给未来自己的一封信」，AI加密保存：\n"
         "  · 设定未来解锁日期（如1年后、3年后）\n"
         "  · 到达解锁日时，AI 将这封信与你的实际成就进行对比，生成一段感人的回顾视频\n"
         "  · 优秀的「预言成真」案例会被推荐到社区，成为最强激励素材")

add_para("④ 梦想接力赛", bold=True)
add_para("AI 设计团队级的「梦想接力」活动：\n"
         "  · 每个团队设定共同目标（如团队总PV达到某个数值）\n"
         "  · AI 生成团队梦想进度的实时可视化动画（如一艘正在航行的帆船）\n"
         "  · 每位成员的贡献实时反映在团队可视化中，增强归属感和责任感")

add_para("⑤ AI 颁奖典礼", bold=True)
add_para("每月/每季度，AI 自动为表现突出的会员举办「虚拟颁奖典礼」：\n"
         "  · AI生成一个逼真的颁奖舞台场景\n"
         "  · 会员的数字分身走上红毯，接受「最佳进步奖」「梦想先锋奖」等荣誉\n"
         "  · 颁奖视频可下载并分享，成为会员最有价值的「数字荣誉勋章」")

doc.add_page_break()

# --- 模块5（新增） ---
add_heading_styled("模块5：🧠 AI 人生导演 —— 智能剧本优化器（全新创意）", level=2, color=RGBColor(0, 102, 153))

add_para("这是提案中最具前瞻性的模块——让AI成为你人生剧本的「导演」。",
         italic=True, color=RGBColor(100, 100, 100))

add_para("① 剧本健康度评分", bold=True)
add_para("AI 对用户的人生剧本进行多维度评估：\n"
         "  · 可行性评分：目标是否合理？时间线是否科学？\n"
         "  · 完整性评分：是否覆盖了人生的各个维度？\n"
         "  · 挑战性评分：目标是太容易了还是不切实际？\n"
         "  · 连贯性评分：不同目标之间是否存在冲突？\n"
         "AI给出综合评分和优化建议，帮助用户写出「最佳剧本」。")

add_para("② 「如果……会怎样」模拟器", bold=True)
add_para("用户可以提出各种假设，AI 进行情景模拟：\n"
         "  · 「如果我每天多花1小时做推广，3个月后我的团队会怎样？」\n"
         "  · 「如果我现在开始学习视频制作，对我的社交媒体运营有多大影响？」\n"
         "  · AI 基于同类型会员的历史数据进行回归分析，给出概率化的预测结果\n"
         "  · 结果以可视化的「平行时间线」形式呈现")

add_para("③ 智能剧本推荐", bold=True)
add_para("基于用户的背景、性格测评和当前状态，AI 推荐「参考剧本」：\n"
         "  · 「和你背景相似的成功会员张三，她的第一年是这样规划的→」\n"
         "  · 「根据你的性格特点（外向型/分析型），建议你采用社交驱动策略」\n"
         "  · 推荐剧本可以一键导入并根据个人情况调整")

add_para("④ 剧本版本管理", bold=True)
add_para("像软件开发一样管理你的人生剧本：\n"
         "  · 保留每一版剧本的历史记录\n"
         "  · AI 分析你的剧本如何演变：「你的梦想从第一版的10个减少到现在的5个，"
         "但每一个都更加聚焦和深入了」\n"
         "  · 生成「人生剧本演化图」，展示你从迷茫到清晰的成长轨迹")

add_para("⑤ AI 人生复盘报告", bold=True)
add_para("每年末，AI 自动生成年度人生复盘报告：\n"
         "  · 对比年初剧本 vs 年末实际成就\n"
         "  · 分析哪些目标超额完成，哪些未达预期及原因\n"
         "  · 基于本年度数据，AI 自动草拟下一年度的剧本建议\n"
         "  · 配以精美的视觉报告和时间轴动画，成为最有仪式感的年终总结")

doc.add_page_break()

# ============ 四、技术架构概览 ============
add_heading_styled("四、技术架构概览", level=1)

add_para("本提案涉及的核心技术栈：")

tech_items = [
    ("生成式视频模型：", "基于Sora/Kling等先进模型，定制化训练，支持数字分身生成"),
    ("AI Agent框架：", "基于LangChain/AutoGPT架构，支持多步骤任务自动执行"),
    ("多模态大语言模型：", "GPT-4o/Claude等，用于自然语言理解、目标拆解和个性化建议"),
    ("3D可视化引擎：", "基于Three.js/Unity WebGL，支持浏览器端3D场景渲染"),
    ("数据分析平台：", "基于实时数据流处理，接入艾多美业务系统API"),
    ("隐私与安全：", "端到端加密、联邦学习、差分隐私等技术确保数据安全"),
]
for prefix, desc in tech_items:
    add_bullet(desc, bold_prefix=prefix)

doc.add_page_break()

# ============ 五、与艾多美业务的深度融合 ============
add_heading_styled("五、与艾多美业务的深度融合", level=1)

add_para("本提案的每一个模块都紧密围绕艾多美的核心业务和企业文化：", bold=True)

fusion_items = [
    ("赋能会员增长：", "AI Agent帮助会员更高效地拓展客户和团队，降低新人上手门槛"),
    ("增强团队凝聚力：", "梦想共振网络让团队成员因共同的梦想而紧密连接"),
    ("强化品牌形象：", "电影级人生预演视频自带传播属性，让每位会员成为品牌大使"),
    ("提升留存率：", "实时进度追踪和情感化反馈机制有效降低会员流失率"),
    ("数据驱动决策：", "AI数据分析帮助公司更精准地了解会员需求和市场趋势"),
    ("践行企业理念：", "「让每个人都能实现梦想」——AI让这句话从口号变成可衡量的现实"),
]
for prefix, desc in fusion_items:
    add_bullet(desc, bold_prefix=prefix)

doc.add_page_break()

# ============ 六、实施路线图 ============
add_heading_styled("六、实施路线图", level=1)

phases = [
    ("第一阶段：MVP验证", "核心功能",
     ["人生剧本文字输入 + AI智能优化建议",
      "基础版「时空穿梭机」（文字→图片→短视频）",
      "基础版梦想仪表盘（进度条+数据看板）",
      "小范围内测（100名种子用户）"]),
    ("第二阶段：功能完善", "体验升级",
     ["数字分身系统上线",
      "AI Agent目标拆解引擎上线",
      "3D梦想星球Beta版",
      "剧本社交基础功能（梦想匹配+时间胶囊）",
      "扩大测试范围（1000名用户）"]),
    ("第三阶段：全面上线", "生态构建",
     ["全功能正式上线",
      "AI人生导演系统完善",
      "多语言支持（覆盖艾多美全球市场）",
      "合作伙伴生态（VR设备、健康数据接入）",
      "持续迭代优化"]),
]

for title, subtitle, items in phases:
    add_para(f"▎{title}——{subtitle}", bold=True, color=RGBColor(0, 102, 153), size=Pt(12))
    for item in items:
        add_bullet(item)

doc.add_page_break()

# ============ 七、预期效果与价值 ============
add_heading_styled("七、预期效果与价值", level=1)

add_heading_styled("7.1 对会员的价值", level=2, color=RGBColor(0, 102, 153))
values_member = [
    "梦想从模糊变为清晰：AI帮助科学地规划人生路径",
    "行动力显著提升：每天有具体的待办事项和AI教练督促",
    "成就感和归属感增强：进度可视化+社交激励机制",
    "个人品牌建设：AI生成的高质量内容提升社交影响力",
]
for v in values_member:
    add_bullet(v)

add_heading_styled("7.2 对公司的价值", level=2, color=RGBColor(0, 102, 153))
values_company = [
    "差异化竞争优势：全球直销行业首创AI人生剧本系统",
    "会员活跃度预计提升50%+",
    "新会员留存率预计提升30%+",
    "社交媒体自然传播量预计增长200%+（用户自发分享AI生成的梦想视频）",
    "数据资产沉淀：海量用户行为和目标数据为公司战略决策提供支撑",
]
for v in values_company:
    add_bullet(v)

add_heading_styled("7.3 行业影响力", level=2, color=RGBColor(0, 102, 153))
add_para("本方案若落地，艾多美将成为全球首家将AI Agent深度融入会员目标管理体系的直销企业，"
         "开创「AI驱动的梦想实现」新范式，具有标杆效应和巨大的媒体传播价值。",
         bold=True)

doc.add_page_break()

# ============ 八、总结 ============
add_heading_styled("八、总结", level=1)

add_para("「书写人生剧本」一直是艾多美最打动人心的理念之一。但在AI时代，"
         "我们可以让这份理念产生质的飞跃——", size=Pt(12))

doc.add_paragraph()

summary_lines = [
    "从「写下来」到「看得见」—— 时空穿梭机让梦想变成电影",
    "从「想一想」到「跑起来」—— AI Agent让目标自动执行",
    "从「凭感觉」到「有数据」—— 梦想仪表盘让进度透明",
    "从「一个人」到「一群人」—— 社交宇宙让梦想共振",
    "从「拍脑袋」到「有智慧」—— AI导演让剧本不断进化",
]
for line in summary_lines:
    add_para(f"✦ {line}", bold=True, color=RGBColor(0, 51, 102), size=Pt(12), space_after=Pt(8))

doc.add_paragraph()

add_para("我们的愿景是：", bold=True, size=Pt(14), color=RGBColor(0, 51, 102))
add_para("让每一位艾多美人都拥有一个AI驱动的「人生操作系统」，", 
         bold=True, size=Pt(14), color=RGBColor(200, 50, 50), align=WD_ALIGN_PARAGRAPH.CENTER)
add_para("不只是书写剧本，更是活在自己最好的剧本里。",
         bold=True, size=Pt(14), color=RGBColor(200, 50, 50), align=WD_ALIGN_PARAGRAPH.CENTER)

doc.add_paragraph()
doc.add_paragraph()
add_para("—— 感谢评审老师的审阅 ——", align=WD_ALIGN_PARAGRAPH.CENTER, 
         color=RGBColor(150, 150, 150), italic=True)

# ============ 保存 ============
output_path = "/workspace/艾多美AI创意提案_AI人生剧本3.0.docx"
doc.save(output_path)
print(f"文档已保存至: {output_path}")
