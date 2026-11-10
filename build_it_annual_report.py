from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import ChartData
from pathlib import Path

ROOT = Path(r"D:\Code\Myagent-System-Skills")
OUT = ROOT / "output"
OUT.mkdir(exist_ok=True)
OUTFILE = OUT / "信息部年终述职报告_虚拟数据.pptx"
BODY_BG = ROOT / "pptx/assets/body-background.jpeg"
COVER_BG = ROOT / "pptx/assets/cover-background.jpeg"

W, H = 10.0, 5.625
FONT = "Microsoft YaHei"
BLUE = "374B9B"
NAVY = "445469"
RED = "E63237"
ORANGE = "ED7D31"
INK = "0F172A"
MUTED = "64748B"
LIGHT = "EEF2F7"
WHITE = "FFFFFF"
GREEN = "1E8E5A"

prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)
blank = prs.slide_layouts[6]

def rgb(hexstr):
    return RGBColor.from_string(hexstr)

def add_bg(slide, cover=False):
    slide.shapes.add_picture(str(COVER_BG if cover else BODY_BG), 0, 0, width=Inches(W), height=Inches(H))

def add_text(slide, text, x, y, w, h, size=12, color=INK, bold=False, align=PP_ALIGN.LEFT,
             valign=MSO_ANCHOR.TOP, margin=0.03, name=None, italic=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.clear()
    tf.margin_left = Inches(margin); tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin); tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(0)
    run = p.add_run(); run.text = text
    run.font.name = FONT; run.font.size = Pt(size); run.font.bold = bold; run.font.italic = italic; run.font.color.rgb = rgb(color)
    return box

def add_title(slide, title, page):
    add_text(slide, title, 1.68, 0.105, 6.8, 0.36, size=22.5, color=WHITE, bold=True, margin=0)
    add_text(slide, f"{page:02d}", 9.03, 0.11, 0.45, 0.25, size=8.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER, margin=0)
    add_text(slide, "信息部｜年度述职", 0.52, 5.335, 2.0, 0.15, size=8.0, color=NAVY, margin=0)
    add_text(slide, f"{page:02d}", 9.25, 5.335, 0.25, 0.15, size=8.0, color=NAVY, bold=True, align=PP_ALIGN.RIGHT, margin=0)

def add_round_rect(slide, x, y, w, h, fill=WHITE, line="D9E1EC", radius=True):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                   Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid(); shape.fill.fore_color.rgb = rgb(fill)
    shape.line.color.rgb = rgb(line); shape.line.width = Pt(0.75)
    return shape

def add_metric(slide, x, y, w, label, value, suffix, accent=BLUE, note=""):
    add_round_rect(slide, x, y, w, 1.18, WHITE, "D9E1EC")
    slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(0.07), Inches(1.18)).fill.solid()
    bar = slide.shapes[-1]; bar.fill.fore_color.rgb = rgb(accent); bar.line.fill.background()
    add_text(slide, label, x+0.18, y+0.16, w-0.3, 0.2, size=10.5, color=MUTED, margin=0)
    add_text(slide, value, x+0.18, y+0.40, w-0.62, 0.42, size=25, color=accent, bold=True, margin=0)
    add_text(slide, suffix, x+w-0.5, y+0.58, 0.35, 0.2, size=10.5, color=INK, bold=True, margin=0)
    if note: add_text(slide, note, x+0.18, y+0.91, w-0.3, 0.16, size=8.5, color=MUTED, margin=0)

def add_bullet(slide, x, y, text, color=INK, accent=BLUE, w=3.9, size=11.5):
    dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y+0.05), Inches(0.12), Inches(0.12))
    dot.fill.solid(); dot.fill.fore_color.rgb = rgb(accent); dot.line.fill.background()
    add_text(slide, text, x+0.2, y, w-0.2, 0.35, size=size, color=color, margin=0)

def add_bar(slide, x, y, label, value, maxv, color=BLUE, w=3.5, suffix=""):
    add_text(slide, label, x, y, 1.1, 0.22, size=10.5, color=INK, margin=0)
    bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x+1.15), Inches(y+0.03), Inches(w), Inches(0.16))
    bg.fill.solid(); bg.fill.fore_color.rgb = rgb("E5EAF2"); bg.line.fill.background()
    fg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x+1.15), Inches(y+0.03), Inches(w*value/maxv), Inches(0.16))
    fg.fill.solid(); fg.fill.fore_color.rgb = rgb(color); fg.line.fill.background()
    add_text(slide, f"{value}{suffix}", x+1.15+w+0.08, y-0.01, 0.7, 0.22, size=10.5, color=INK, bold=True, margin=0)

def add_callout(slide, x, y, w, h, title, body, accent=BLUE):
    add_round_rect(slide, x, y, w, h, "F7F9FC", "D9E1EC")
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(0.06), Inches(h))
    line.fill.solid(); line.fill.fore_color.rgb = rgb(accent); line.line.fill.background()
    add_text(slide, title, x+0.18, y+0.15, w-0.28, 0.24, size=12.5, color=accent, bold=True, margin=0)
    add_text(slide, body, x+0.18, y+0.47, w-0.3, h-0.56, size=10.5, color=INK, margin=0)

# 1 cover
slide = prs.slides.add_slide(blank); add_bg(slide, True)
add_text(slide, "信息部年终述职报告", 0.85, 1.55, 8.3, 0.55, size=28.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER, margin=0)
add_text(slide, "以数字化底座支撑业务增长｜2026 年度复盘与 2027 年规划", 0.86, 2.62, 8.25, 0.28, size=14, color="E8ECF8", align=PP_ALIGN.CENTER, margin=0)
add_text(slide, "汇报人：信息部  ｜  日期：2026年12月", 5.20, 4.86, 4.20, 0.25, size=8.8, color=NAVY, align=PP_ALIGN.RIGHT, margin=0)
add_text(slide, "虚拟数据演示稿｜可按实际经营数据替换", 0.86, 4.86, 4.0, 0.25, size=8.5, color="E8ECF8", margin=0)

# 2 toc
slide = prs.slides.add_slide(blank); add_bg(slide); add_title(slide, "目录｜从结果复盘到下一步行动", 2)
items = [("01", "年度总览", "核心指标与年度判断"), ("02", "重点项目", "业务系统与平台交付"), ("03", "运营提效", "数据、服务与安全治理"), ("04", "团队与规划", "组织能力与 2027 重点")]
for i,(n,t,d) in enumerate(items):
    x = 0.72 + (i%2)*4.42; y = 1.20 + (i//2)*1.55
    add_round_rect(slide, x, y, 3.95, 1.08, WHITE, "D9E1EC")
    add_text(slide, n, x+0.18, y+0.18, 0.55, 0.32, size=22, color=BLUE if i<3 else RED, bold=True, margin=0)
    add_text(slide, t, x+0.86, y+0.20, 2.8, 0.25, size=15, color=INK, bold=True, margin=0)
    add_text(slide, d, x+0.86, y+0.56, 2.8, 0.20, size=10.5, color=MUTED, margin=0)
add_callout(slide, 0.72, 4.38, 8.45, 0.62, "年度判断", "数字化投入从“项目交付”转向“业务结果”：稳定性、效率与风险控制成为同一张成绩单。", RED)

# 3 overview
slide = prs.slides.add_slide(blank); add_bg(slide); add_title(slide, "年度总览｜数字化底座持续增强，业务体验显著改善", 3)
add_text(slide, "2026 年信息部围绕“稳运行、提效率、控风险”三条主线，完成 4 项关键跃迁。", 0.62, 0.78, 8.5, 0.32, size=13, color=INK, bold=True, margin=0)
metrics = [("系统可用性", "99.96", "%", BLUE, "较上年 +0.18pp"), ("项目按期交付", "92", "%", GREEN, "12 个重点项目"), ("人均工单处理", "41", "单/周", ORANGE, "较上年 +24%"), ("安全事件", "0", "起", RED, "重大事件为 0")]
for i,m in enumerate(metrics): add_metric(slide, 0.62+i*2.27, 1.28, 2.08, *m)
add_round_rect(slide, 0.62, 2.75, 4.25, 1.85, "F7F9FC", "D9E1EC")
add_text(slide, "年度关键动作", 0.82, 2.93, 3.8, 0.24, size=13.5, color=BLUE, bold=True, margin=0)
for j,t in enumerate(["统一监控口径，建立 24×7 值守机制", "把高频流程搬到低代码与自动化平台", "以分级保护和演练闭环压降安全风险"]): add_bullet(slide, 0.84, 3.32+j*0.37, t, w=3.75)
add_round_rect(slide, 5.12, 2.75, 4.05, 1.85, "F7F9FC", "D9E1EC")
add_text(slide, "年度复盘结论", 5.32, 2.93, 3.65, 0.24, size=13.5, color=RED, bold=True, margin=0)
add_text(slide, "从“交付 IT 项目”走向“交付业务确定性”。", 5.32, 3.34, 3.45, 0.42, size=17, color=INK, bold=True, margin=0)
add_text(slide, "稳定性与效率指标已进入经营例会，信息部开始用同一套数据语言与业务对话。", 5.32, 3.92, 3.45, 0.45, size=10.5, color=MUTED, margin=0)

# 4 key projects
slide = prs.slides.add_slide(blank); add_bg(slide); add_title(slide, "重点项目｜四个关键项目把“效率提升”落到业务现场", 4)
projects = [
    ("经营驾驶舱", "把 8 张经营报表整合成一张实时看板", "交付：Q1–Q2", "业务决策时效 -40%", BLUE),
    ("协同办公升级", "统一流程、权限与移动端体验", "交付：Q2", "审批平均时长 -32%", GREEN),
    ("供应链计划平台", "打通计划、库存、采购三类数据", "交付：Q2–Q3", "库存周转 +12%", ORANGE),
    ("主数据治理", "建立客户、物料、组织三大主数据域", "交付：Q3–Q4", "数据重复率 -68%", RED),
]
for i,(t,b,tag,result,accent) in enumerate(projects):
    x=0.62+(i%2)*4.42; y=1.08+(i//2)*1.93
    add_round_rect(slide,x,y,4.05,1.55,WHITE,"D9E1EC")
    add_text(slide,f"0{i+1}",x+0.18,y+0.18,0.42,0.3,size=18,color=accent,bold=True,margin=0)
    add_text(slide,t,x+0.75,y+0.16,3.05,0.25,size=14,color=INK,bold=True,margin=0)
    add_text(slide,b,x+0.75,y+0.52,3.0,0.36,size=10.5,color=MUTED,margin=0)
    add_text(slide,tag,x+0.75,y+0.99,1.2,0.2,size=9,color=accent,bold=True,margin=0)
    add_text(slide,result,x+2.0,y+0.99,1.7,0.2,size=10.5,color=INK,bold=True,align=PP_ALIGN.RIGHT,margin=0)
add_callout(slide,0.62,4.72,8.45,0.48,"交付方法", "每个项目绑定业务负责人、结果指标和上线后 30 天复盘，确保“上线”不等于“结束”。", BLUE)

# 5 platform/data
slide = prs.slides.add_slide(blank); add_bg(slide); add_title(slide, "平台与数据｜从“可用”到“可复用”，平台能力开始沉淀", 5)
add_text(slide, "平台化建设让一次性交付变成可复制的能力组件。", 0.62, 0.80, 8.2, 0.3, size=13, color=INK, bold=True, margin=0)
add_round_rect(slide,0.62,1.30,5.2,3.45,"F7F9FC","D9E1EC")
add_text(slide,"平台能力成熟度（虚拟评分）",0.85,1.50,4.6,0.24,size=13.5,color=BLUE,bold=True,margin=0)
labels=[("统一身份与权限",86,BLUE),("数据服务目录",78,GREEN),("自动化流程组件",72,ORANGE),("监控与告警",91,RED)]
for i,(l,v,c) in enumerate(labels): add_bar(slide,0.88,1.98+i*0.52,l,v,100,c,w=2.75,suffix="分")
add_text(slide,"评分口径：覆盖率、复用率、可观测性与业务满意度综合评分",0.88,4.32,4.6,0.2,size=8.5,color=MUTED,margin=0)
add_round_rect(slide,6.12,1.30,3.0,3.45,WHITE,"D9E1EC")
add_text(slide,"数据资产沉淀",6.36,1.50,2.5,0.24,size=13.5,color=RED,bold=True,margin=0)
for i,(num,lab,sub) in enumerate([("126", "数据集市表", "张"),("38", "标准 API", "个"),("74%", "主数据覆盖", "率")]):
    yy=1.98+i*0.82
    add_text(slide,num,6.36,yy,0.95,0.34,size=23,color=RED if i==0 else BLUE,bold=True,margin=0)
    add_text(slide,lab,7.42,yy+0.02,1.35,0.22,size=11,color=INK,bold=True,margin=0)
    add_text(slide,sub,7.42,yy+0.31,1.2,0.17,size=8.5,color=MUTED,margin=0)
add_callout(slide,6.12,4.98,3.0,0.42,"下一步", "把数据产品纳入业务产品路线图", ORANGE)

# 6 security
slide = prs.slides.add_slide(blank); add_bg(slide); add_title(slide, "安全合规｜用可视化、可演练、可追责守住底线", 6)
add_round_rect(slide,0.62,1.02,3.0,3.95,"F7F9FC","D9E1EC")
add_text(slide,"安全运营闭环",0.86,1.25,2.4,0.24,size=14,color=BLUE,bold=True,margin=0)
steps=[("01","识别","资产盘点 100%"),("02","防护","高危漏洞 48h 内闭环"),("03","演练","全年 6 次专项演练"),("04","复盘","问题归因与责任到人")]
for i,(n,t,d) in enumerate(steps):
    yy=1.78+i*0.70
    circ=slide.shapes.add_shape(MSO_SHAPE.OVAL,Inches(0.88),Inches(yy),Inches(0.36),Inches(0.36)); circ.fill.solid(); circ.fill.fore_color.rgb=rgb(BLUE); circ.line.fill.background()
    add_text(slide,n,0.88,yy+0.07,0.36,0.16,size=8.5,color=WHITE,bold=True,align=PP_ALIGN.CENTER,margin=0)
    add_text(slide,t,1.42,yy+0.02,1.1,0.20,size=11.5,color=INK,bold=True,margin=0)
    add_text(slide,d,1.42,yy+0.27,1.75,0.18,size=9.5,color=MUTED,margin=0)
    if i<3:
        line=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(1.05),Inches(yy+0.36),Inches(0.02),Inches(0.34)); line.fill.solid(); line.fill.fore_color.rgb=rgb("B9C4D8"); line.line.fill.background()
add_round_rect(slide,3.95,1.02,5.17,1.70,WHITE,"D9E1EC")
add_text(slide,"风险趋势",4.20,1.23,2.0,0.22,size=13.5,color=RED,bold=True,margin=0)
add_text(slide,"高危漏洞平均修复时长",4.20,1.70,2.1,0.18,size=10.5,color=MUTED,margin=0)
add_text(slide,"72h → 31h",4.20,1.93,2.2,0.34,size=24,color=RED,bold=True,margin=0)
add_text(slide,"下降 57%",7.58,1.99,1.0,0.2,size=11,color=GREEN,bold=True,align=PP_ALIGN.RIGHT,margin=0)
add_round_rect(slide,3.95,3.00,5.17,1.97,"F7F9FC","D9E1EC")
add_text(slide,"年度安全结果",4.20,3.22,2.2,0.22,size=13.5,color=BLUE,bold=True,margin=0)
for i,(lab,val,col) in enumerate([("重大安全事件","0 起",GREEN),("关键系统纳管率","100%",BLUE),("安全培训覆盖","98%",ORANGE)]):
    xx=4.20+i*1.55; add_text(slide,val,xx,3.76,1.3,0.30,size=19,color=col,bold=True,margin=0); add_text(slide,lab,xx,4.16,1.35,0.20,size=9.5,color=MUTED,margin=0)

# 7 operations
slide = prs.slides.add_slide(blank); add_bg(slide); add_title(slide, "运营提效｜把服务台从“接单中心”升级为“体验中心”", 7)
add_text(slide,"高频问题自动化、重复问题知识化、重大问题闭环化，带动服务体验连续改善。",0.62,0.80,8.5,0.3,size=13,color=INK,bold=True,margin=0)
add_round_rect(slide,0.62,1.28,5.25,3.40,"F7F9FC","D9E1EC")
add_text(slide,"服务指标趋势（虚拟数据）",0.88,1.48,3.5,0.23,size=13.5,color=BLUE,bold=True,margin=0)
add_text(slide,"指标",0.88,1.92,0.7,0.18,size=9.5,color=MUTED,margin=0); add_text(slide,"2025",3.10,1.92,0.7,0.18,size=9.5,color=MUTED,align=PP_ALIGN.RIGHT,margin=0); add_text(slide,"2026",4.52,1.92,0.7,0.18,size=9.5,color=MUTED,align=PP_ALIGN.RIGHT,margin=0)
rows=[("首响时长（小时）",4.8,2.1,RED),("一次解决率",68,84,GREEN),("满意度",88,95,BLUE),("自助解决率",31,57,ORANGE)]
for i,(lab,a,b,col) in enumerate(rows):
    yy=2.28+i*0.50; add_text(slide,lab,0.88,yy,1.75,0.2,size=10.5,color=INK,margin=0)
    # bars in two columns
    bg1=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(2.70),Inches(yy+0.03),Inches(1.05),Inches(0.14)); bg1.fill.solid(); bg1.fill.fore_color.rgb=rgb("E5EAF2"); bg1.line.fill.background()
    bg2=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.12),Inches(yy+0.03),Inches(1.05),Inches(0.14)); bg2.fill.solid(); bg2.fill.fore_color.rgb=rgb("E5EAF2"); bg2.line.fill.background()
    maxv=max(a,b,100 if isinstance(a,(int,float)) and a>90 else 100)
    f1=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(2.70),Inches(yy+0.03),Inches(1.05*(a/maxv)),Inches(0.14)); f1.fill.solid(); f1.fill.fore_color.rgb=rgb("A9B8D6"); f1.line.fill.background()
    f2=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.12),Inches(yy+0.03),Inches(1.05*(b/maxv)),Inches(0.14)); f2.fill.solid(); f2.fill.fore_color.rgb=rgb(col); f2.line.fill.background()
    add_text(slide,str(a),3.83,yy-0.01,0.35,0.18,size=9.5,color=MUTED,align=PP_ALIGN.RIGHT,margin=0); add_text(slide,str(b),5.25,yy-0.01,0.35,0.18,size=9.5,color=col,bold=True,align=PP_ALIGN.RIGHT,margin=0)
add_text(slide,"单位：小时 / %；数据为内部演示口径",0.88,4.38,3.8,0.18,size=8.5,color=MUTED,margin=0)
add_round_rect(slide,6.16,1.28,2.96,3.40,WHITE,"D9E1EC")
add_text(slide,"三个可复制动作",6.42,1.48,2.4,0.23,size=13.5,color=RED,bold=True,margin=0)
for i,(t,d) in enumerate([("知识库前置", "把 20 个高频问题变成自助入口"),("机器人分流", "简单问题自动处理，专家聚焦复杂问题"),("月度复盘", "用数据推动服务目录持续优化")]):
    yy=1.98+i*0.80; add_text(slide,f"0{i+1}",6.42,yy,0.35,0.24,size=14,color=RED,bold=True,margin=0); add_text(slide,t,6.92,yy,1.7,0.2,size=11.5,color=INK,bold=True,margin=0); add_text(slide,d,6.92,yy+0.28,1.85,0.34,size=9.5,color=MUTED,margin=0)

# 8 team
slide = prs.slides.add_slide(blank); add_bg(slide); add_title(slide, "团队建设｜用机制把个人经验变成组织能力", 8)
add_text(slide,"团队规模保持稳定，能力结构从“系统运维”向“产品、数据、安全”三栖演进。",0.62,0.80,8.5,0.3,size=13,color=INK,bold=True,margin=0)
add_metric(slide,0.62,1.35,2.1,"团队规模","18","人",BLUE,"研发 / 运维 / 数据 / 安全")
add_metric(slide,2.94,1.35,2.1,"认证与训练","26","人次",GREEN,"云、数据、安全方向")
add_metric(slide,5.26,1.35,2.1,"知识沉淀","74","篇",ORANGE,"标准作业与复盘文档")
add_metric(slide,7.58,1.35,1.54,"满意度","92","分",RED,"内部调研")
add_round_rect(slide,0.62,2.85,4.25,1.85,"F7F9FC","D9E1EC")
add_text(slide,"组织能力三件事",0.86,3.05,3.6,0.23,size=13.5,color=BLUE,bold=True,margin=0)
for i,t in enumerate(["岗位职责与服务目录一一对应", "关键系统建立双人备份与轮值机制", "项目复盘沉淀为可复用模板"]): add_bullet(slide,0.88,3.45+i*0.35,t,w=3.75,size=10.5)
add_round_rect(slide,5.12,2.85,4.0,1.85,WHITE,"D9E1EC")
add_text(slide,"团队共识",5.36,3.05,3.2,0.23,size=13.5,color=RED,bold=True,margin=0)
add_text(slide,"稳定运行是底线\n业务结果是标准\n持续学习是方法",5.36,3.48,3.2,0.82,size=16,color=INK,bold=True,margin=0)

# 9 plan
slide = prs.slides.add_slide(blank); add_bg(slide); add_title(slide, "2027 规划｜围绕“平台化、智能化、韧性化”再上一个台阶", 9)
add_text(slide,"明年重点从“把事情做好”升级为“让组织更快、更稳、更聪明”。",0.62,0.80,8.5,0.3,size=13,color=INK,bold=True,margin=0)
plans=[("01","平台化", "统一技术底座", "完成身份、数据、流程三类平台的统一入口与服务目录", BLUE), ("02","智能化", "AI 进入业务流程", "在客服、经营分析、知识管理三类场景形成可量化收益", ORANGE), ("03","韧性化", "从恢复走向预防", "建立关键系统容灾分级与季度演练机制，目标 RTO ≤ 30 分钟", RED)]
for i,(n,t,sub,body,col) in enumerate(plans):
    x=0.62+i*2.83; add_round_rect(slide,x,1.34,2.55,2.35,WHITE,"D9E1EC")
    add_text(slide,n,x+0.18,1.58,0.48,0.28,size=18,color=col,bold=True,margin=0)
    add_text(slide,t,x+0.78,1.59,1.45,0.22,size=14,color=INK,bold=True,margin=0)
    add_text(slide,sub,x+0.18,2.06,2.05,0.22,size=11.5,color=col,bold=True,margin=0)
    add_text(slide,body,x+0.18,2.48,2.12,0.72,size=10.2,color=MUTED,margin=0)
add_round_rect(slide,0.62,4.05,8.45,0.70,"F7F9FC","D9E1EC")
add_text(slide,"2027 目标",0.88,4.26,1.0,0.22,size=13.5,color=BLUE,bold=True,margin=0)
add_text(slide,"业务系统可用性 ≥ 99.98%   ｜   自动化流程覆盖率 ≥ 60%   ｜   重大安全事件 = 0",2.05,4.24,6.5,0.24,size=12.5,color=INK,bold=True,margin=0)

# 10 end
slide = prs.slides.add_slide(blank); add_bg(slide, True)
add_text(slide,"谢谢聆听",0.85,1.62,8.3,0.55,size=28.5,color=WHITE,bold=True,align=PP_ALIGN.CENTER,margin=0)
add_text(slide,"让技术成为业务确定性的底座",0.86,2.62,8.25,0.28,size=14,color="E8ECF8",align=PP_ALIGN.CENTER,margin=0)
add_text(slide,"信息部  ｜  2026 年终述职",5.20,4.86,4.20,0.25,size=8.8,color=NAVY,align=PP_ALIGN.RIGHT,margin=0)

# document metadata
prs.core_properties.title = "信息部年终述职报告（虚拟数据）"
prs.core_properties.subject = "2026年度信息部工作复盘与2027年规划"
prs.core_properties.author = "信息部"
prs.core_properties.comments = "本演示文稿全部业务数据为虚拟演示数据，请按实际情况替换。"
prs.save(OUTFILE)
print(OUTFILE)
