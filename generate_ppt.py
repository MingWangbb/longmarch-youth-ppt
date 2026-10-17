from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

# ---------------------------
# 配置
# ---------------------------
OUTPUT_FILE = "长征路上的青春选择.pptx"
REPORTER = "[姓名]"
DATE = "202X年X月X日"
EVENT_NAME = "[课程/活动名称]"

# 颜色
RED = RGBColor(105, 14, 18)
DEEP_RED = RGBColor(56, 10, 16)
GOLD = RGBColor(212, 175, 55)
DARK_GOLD = RGBColor(174, 131, 53)
WHITE = RGBColor(255, 255, 255)
BLACK = RGBColor(0, 0, 0)
GRAY = RGBColor(70, 76, 82)
LIGHT_GRAY = RGBColor(236, 236, 236)
BLUE = RGBColor(22, 62, 91)
BLUE_LIGHT = RGBColor(117, 146, 167)
MINT = RGBColor(172, 184, 175)

# ---------------------------
# 基础工具
# ---------------------------
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)


def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, left, top, width, height, text, font_size=18, bold=False, color=BLACK,
                align=PP_ALIGN.LEFT, font_name="Microsoft YaHei", line_spacing=1.2, margin=0.08):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    p = tf.paragraphs[0]
    p.alignment = align
    p.text = text
    p.space_after = Pt(0)
    p.line_spacing = line_spacing
    run = p.runs[0]
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = font_name
    run.font.color.rgb = color
    return box


def add_footer(slide, text):
    add_textbox(slide, Inches(0.8), Inches(6.95), Inches(12.0), Inches(0.35), text,
                font_size=11, color=WHITE, bold=False, align=PP_ALIGN.LEFT, font_name="Microsoft YaHei")


def add_bullets(slide, left, top, width, height, items, font_size=20, color=BLACK, indent=0.25):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.12)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.06)
    tf.margin_bottom = Inches(0.06)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.bullet = True
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(8)
        p.line_spacing = 1.2
        run = p.runs[0]
        run.font.size = Pt(font_size)
        run.font.name = "Microsoft YaHei"
        run.font.color.rgb = color
        run.font.bold = False

# ---------------------------
# Slide 1: 封面
# ---------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, RGBColor(17, 18, 19))

left_panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.55), Inches(0.95), Inches(5.1), Inches(4.8))
left_panel.fill.solid()
left_panel.fill.fore_color.rgb = RGBColor(86, 86, 86)
left_panel.line.color.rgb = RGBColor(86, 86, 86)

right_panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.6), Inches(1.1), Inches(5.0), Inches(4.5))
right_panel.fill.solid()
right_panel.fill.fore_color.rgb = RGBColor(110, 110, 110)
right_panel.line.color.rgb = RGBColor(110, 110, 110)

band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.65), Inches(1.2), Inches(2.1), Inches(4.2))
band.fill.solid()
band.fill.fore_color.rgb = RGBColor(220, 183, 90)
band.line.color.rgb = RGBColor(220, 183, 90)

add_textbox(slide, Inches(1.8), Inches(1.7), Inches(9.8), Inches(1.5),
            "长征路上的青春选择：\n在历史回响中锚定人生坐标",
            font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(slide, Inches(2.2), Inches(3.2), Inches(9.0), Inches(1.0),
            "领悟人生真谛·把握人生方向·担当复兴大任·成就时代新人",
            font_size=16, bold=False, color=RGBColor(232, 220, 190), align=PP_ALIGN.CENTER)
add_footer(slide, f"汇报人：{REPORTER} | {DATE} | {EVENT_NAME}")

# ---------------------------
# Slide 2: 目录页
# ---------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, RGBColor(26, 30, 35))
add_textbox(slide, Inches(0.75), Inches(0.55), Inches(10.0), Inches(0.5),
            "叙事脉络与思想进阶", font_size=24, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

axis = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(1.1), Inches(11.7), Inches(0.08))
axis.fill.solid(); axis.fill.fore_color.rgb = GOLD; axis.line.color.rgb = GOLD

chapters = [
    ("历史纵深", "生死荣辱间的信仰抉择与价值重塑", 0.9, 1.5, 2.7, 1.5),
    ("精神解码", "长征精神的青春内核与当代转译", 3.9, 1.5, 2.7, 1.5),
    ("现实观照", "内卷、躺平与迷茫背后的时代症候", 6.9, 1.5, 2.7, 1.5),
    ("行动纲领", "以长期主义走好新时代青春长征路", 9.9, 1.5, 2.7, 1.5),
]
for i, (title, subtitle, x, y, w, h) in enumerate(chapters):
    chapter_color = [RGBColor(131, 21, 31), RGBColor(58, 67, 77), RGBColor(19, 83, 89), RGBColor(142, 107, 52)]
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    rect.fill.solid(); rect.fill.fore_color.rgb = chapter_color[i]; rect.line.color.rgb = chapter_color[i]
    add_textbox(slide, Inches(x+0.15), Inches(y+0.18), Inches(w-0.3), Inches(0.6), title, font_size=17, bold=True, color=WHITE)
    add_textbox(slide, Inches(x+0.15), Inches(y+0.75), Inches(w-0.3), Inches(0.55), subtitle, font_size=10, bold=False, color=WHITE)

# 其他页简略示例（省略部分内容）
# 下方为可直接生成12页的基础版本，保留关键页面结构。若需要进一步精修版，请运行后续增强脚本。

# ---------------------------
# Slide 3: 第一章过渡页
# ---------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, RGBColor(56, 12, 16))
bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.65), Inches(12.2), Inches(0.12))
bar.fill.solid(); bar.fill.fore_color.rgb = GOLD; bar.line.color.rgb = GOLD
add_textbox(slide, Inches(0.9), Inches(1.1), Inches(9.0), Inches(0.6),
            "01 历史纵深：绝境中的信仰试炼", font_size=24, bold=True, color=RGBColor(250, 220, 150))
add_textbox(slide, Inches(0.9), Inches(2.3), Inches(10.4), Inches(1.3),
            "“每一代人有每一代人的长征路，每一代人都要走好自己的长征路。”",
            font_size=22, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

# ---------------------------
# Slide 4: 历史的抉择（生死与顺逆）
# ---------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, RGBColor(245, 243, 238))
add_textbox(slide, Inches(0.75), Inches(0.5), Inches(8.5), Inches(0.6),
            "向死而生：在极端困境中重构价值排序", font_size=28, bold=True, color=RGBColor(52, 38, 38))
paragraphs = [
    "生死关头的信仰具象化：湘江战役中，红三十四师作为后卫部队全员牺牲，师长陈树湘断肠明志。这种“虽死犹生”的选择，并非对生命的漠视，而是将个体生命融入民族解放的宏大叙事——当个人存亡与集体命运绑定，死亡便获得了超越性的意义。这启示我们：真正的安全感不来自规避风险，而来自对更高价值的确认。",
    "逆境中的认知迭代能力：从第五次反“围剿”失败到遵义会议确立正确路线，红军经历了从“教条主义”到“实事求是”的痛苦转型。这种在低谷中主动修正认知、打破路径依赖的能力，正是当代青年面对学业挫折、职业转型时最稀缺的素养——逆境不是终点，而是重新校准方向的契机。"
]
add_bullets(slide, Inches(0.9), Inches(1.4), Inches(11.2), Inches(4.5), paragraphs, font_size=15, color=RGBColor(36, 36, 36))

# ---------------------------
# Slide 5: 历史的抉择（苦乐与得失）
# ---------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, RGBColor(249, 247, 242))
add_textbox(slide, Inches(0.75), Inches(0.5), Inches(8.5), Inches(0.6),
            "苦难美学：在物质匮乏中构建精神丰盈", font_size=28, bold=True, color=RGBColor(100, 40, 35))
paragraphs = [
    "苦乐辩证法的实践样本：过草地时，战士们以野菜充饥却高唱《国际歌》；翻雪山时，冻僵的遗体仍保持前进姿态。这种“以苦为乐”并非自我麻痹，而是基于对革命必胜信念的深度认同——当精神目标足够清晰，肉体的痛苦便被转化为通往理想的必经仪式。对比当下部分青年“稍遇压力即崩溃”的现象，本质是缺乏支撑苦难的意义框架。",
    "得失观的时空尺度拓展：红军放弃瑞金舒适区、舍弃个人财产甚至家庭团聚，看似“失去一切”，却在历史长河中赢得了民族新生。这种选择揭示了“得与失”的动态转化规律：短期利益的让渡，往往是为了获取更长远、更根本的价值回报。对当代人而言，拒绝“精致利己主义”，方能跳出零和博弈的内耗陷阱。"
]
add_bullets(slide, Inches(0.9), Inches(1.5), Inches(11.2), Inches(4.5), paragraphs, font_size=15, color=RGBColor(35, 35, 35))

# ---------------------------
# Slide 6: 第二章过渡页
# ---------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, RGBColor(107, 131, 150))
add_textbox(slide, Inches(0.9), Inches(1.0), Inches(9.0), Inches(0.6),
            "02 现实观照：当精神遗产遭遇时代症候", font_size=24, bold=True, color=WHITE)
add_textbox(slide, Inches(0.9), Inches(2.6), Inches(10.0), Inches(1.0),
            "“没有枪林弹雨的和平年代，为何我们反而陷入更深的焦虑？”",
            font_size=24, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

# ---------------------------
# Slide 7: 当代大学生的“新长征”困境
# ---------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, RGBColor(237, 236, 235))
add_textbox(slide, Inches(0.75), Inches(0.55), Inches(8.5), Inches(0.6),
            "三重困境：被异化的成长坐标系", font_size=28, bold=True, color=RGBColor(26, 30, 35))
paragraphs = [
    "迷茫：意义系统的真空化：当“考上大学”这一单一目标达成后，许多新生陷入“空心病”——不知为何学习、为谁奋斗。这如同长征初期失去中央苏区依托后的方向迷失，根源在于未建立与国家发展、人类进步相联结的深层动机。随波逐流考研、考公，实则是用战术勤奋掩盖战略懒惰。",
    "内卷：竞争逻辑的自我殖民：绩点通胀、综测刷分、实习注水……我们在量化指标中疲于奔命，却遗忘了知识探索的本真乐趣。这类似于资源匮乏时期的过度消耗战，将同学视为对手而非伙伴，最终导致集体创造力萎缩。真正的竞争力从不来自同质化厮杀，而源于差异化价值的创造。",
    "躺平：主体性的消极退场：“努力无用论”“佛系摆烂”表面是情绪宣泄，实质是对责任的逃避。就像行军中因畏惧艰险而掉队的士兵，他们误将“休息”等同于“放弃”，却不知长征的胜利恰恰属于那些“走得慢但从未停步”的人。允许阶段性调整，但绝不接受永久性退出。"
]
add_bullets(slide, Inches(0.7), Inches(1.5), Inches(12.0), Inches(4.8), paragraphs, font_size=12, color=RGBColor(20, 20, 20))

# ---------------------------
# Slide 8: 第三章过渡页
# ---------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, RGBColor(236, 240, 234))
add_textbox(slide, Inches(0.9), Inches(0.9), Inches(9.0), Inches(0.6),
            "03 行动纲领：以长征精神重铸青春主体性", font_size=24, bold=True, color=RGBColor(214, 145, 44))
add_textbox(slide, Inches(0.9), Inches(2.7), Inches(10.5), Inches(1.1),
            "“心中有信仰，脚下有力量；眼中有光芒，肩上有担当。”",
            font_size=22, bold=True, color=RGBColor(26, 30, 35), align=PP_ALIGN.LEFT)

# ---------------------------
# Slide 9: 从长征精神中找答案（一）
# ---------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, RGBColor(249, 247, 242))
add_textbox(slide, Inches(0.75), Inches(0.5), Inches(8.5), Inches(0.6),
            "锚定“心中的延安”：构建多层次目标体系", font_size=28, bold=True, color=RGBColor(64, 45, 30))
paragraphs = [
    "宏观层面：将个体生命嵌入国家叙事：红军北上抗日的战略目标，赋予每一步跋涉以历史正当性。当代青年需主动对接国家战略需求（如科技自立自强、乡村振兴、文化传承），在服务大局中定位个人价值——这不是道德绑架，而是获得持久动力的源泉。当你的工作与时代命题共振，倦怠感自然消解。",
    "微观层面：设计可执行的“里程碑计划”：长征胜利由无数个“小目标”累积而成（如突破四道封锁线、强渡乌江）。拒绝好高骛远，将远大理想拆解为学期级、月度级甚至周度级的具体任务（例：本学期精读5本专业经典、完成1次田野调查、掌握1项数据分析工具）。每完成一个小目标，都是对信心的正向强化。"
]
add_bullets(slide, Inches(1.0), Inches(1.8), Inches(10.8), Inches(4.0), paragraphs, font_size=14, color=RGBColor(25, 25, 25))

# ---------------------------
# Slide 10: 从长征精神中找答案（二）
# ---------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, RGBColor(242, 240, 234))
add_textbox(slide, Inches(0.75), Inches(0.5), Inches(8.5), Inches(0.6),
            "掌握“行军的节奏”：在动态平衡中持续进化", font_size=28, bold=True, color=RGBColor(46, 49, 71))
paragraphs = [
    "对抗内卷：以差异化破局同质化竞争：四渡赤水的精髓在于“避实击虚、灵活机动”。与其在热门赛道挤独木桥，不如挖掘自身独特优势（如跨学科思维、实践能力、沟通天赋），开辟属于自己的“根据地”。真正的卓越不是比别人更快，而是走出不可替代的路径。",
    "对抗躺平：以韧性替代爆发力：长征是一场耐力赛，而非短跑冲刺。允许自己偶尔疲惫、短暂休整，但必须守住“继续前行”的底线。建立支持系统（导师、挚友、兴趣社群），在低谷期汲取外部能量；培养成长型思维，将失败视为反馈而非判决。记住：走得慢不可怕，可怕的是停止行走。"
]
add_bullets(slide, Inches(0.9), Inches(1.6), Inches(11.3), Inches(4.4), paragraphs, font_size=14, color=RGBColor(30, 30, 30))

# ---------------------------
# Slide 11: 总结与升华
# ---------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, RGBColor(27, 32, 37))
add_textbox(slide, Inches(0.8), Inches(0.6), Inches(8.5), Inches(0.7),
            "接棒时刻：以青春之我续写民族史诗", font_size=28, bold=True, color=WHITE)
paragraphs = [
    "80多年前，一群平均年龄不足30岁的青年，用血肉之躯踏出人类精神的高峰。他们的选择，定义了何为“大写的人”。",
    "今日之中国，无需我们再历枪林弹雨，但科技创新的“娄山关”、基层治理的“腊子口”、文化自信的“大渡河”依然横亘在前。我们的长征，是在实验室里攻克卡脖子技术，在田间地头书写振兴篇章，在国际舞台讲述中国故事。",
    "人生方向公式：坚定信仰（解决“为何出发”）+ 务实行动（解决“如何抵达”）+ 乐观心态（解决“怎样坚持”）= 无愧时代的青春答卷"
]
add_bullets(slide, Inches(1.1), Inches(1.8), Inches(5.0), Inches(4.0), paragraphs, font_size=14, color=WHITE)

# ---------------------------
# Slide 12: 结束页
# ---------------------------
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, RGBColor(21, 26, 27))
add_textbox(slide, Inches(1.4), Inches(1.3), Inches(10.4), Inches(0.8),
            "走好我们这一代人的长征路！", font_size=30, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER)
add_textbox(slide, Inches(1.6), Inches(2.3), Inches(10.0), Inches(0.8),
            "以信仰之光照亮前行之路，以奋斗之姿回应时代之问", font_size=18, bold=False, color=RGBColor(226, 212, 162),
            align=PP_ALIGN.CENTER)
add_footer(slide, "感谢聆听 | Q&A环节 | 联系方式/二维码（可选）")

# ---------------------------
# 保存
# ---------------------------
prs.save(OUTPUT_FILE)
print(f"已生成 PPT: {OUTPUT_FILE}")
