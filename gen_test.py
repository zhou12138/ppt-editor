"""生成测试用 PPTX"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()

# Slide 1: 标题页
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "Q2 季度报告"
slide.placeholders[1].text = "2026年6月 · 产品团队"

# Slide 2: 内容页
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "核心指标"
body = slide.placeholders[1]
tf = body.text_frame
tf.text = "营收增长 35%"
p = tf.add_paragraph()
p.text = "用户数突破 100 万"
p = tf.add_paragraph()
p.text = "NPS 评分 72"

# Slide 3: 带表格
slide = prs.slides.add_slide(prs.slide_layouts[5])
slide.shapes.title.text = "部门对比"
table = slide.shapes.add_table(3, 3, Inches(1.5), Inches(2), Inches(7), Inches(2.5)).table
table.cell(0,0).text = "部门"
table.cell(0,1).text = "Q1"
table.cell(0,2).text = "Q2"
table.cell(1,0).text = "产品"
table.cell(1,1).text = "120万"
table.cell(1,2).text = "158万"
table.cell(2,0).text = "运营"
table.cell(2,1).text = "80万"
table.cell(2,2).text = "95万"

# Slide 4: 带文本框
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
box1 = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(4), Inches(1))
box1.text_frame.text = "左上角的蓝色框"
for run in box1.text_frame.paragraphs[0].runs:
    run.font.color.rgb = RGBColor(0x00, 0x00, 0xFF)
    run.font.size = Pt(18)

box2 = slide.shapes.add_textbox(Inches(5.5), Inches(0.5), Inches(4), Inches(1))
box2.text_frame.text = "右上角的红色框"
for run in box2.text_frame.paragraphs[0].runs:
    run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    run.font.size = Pt(18)

box3 = slide.shapes.add_textbox(Inches(3), Inches(4), Inches(4), Inches(1.5))
box3.text_frame.text = "总结：本季度表现优异"
for run in box3.text_frame.paragraphs[0].runs:
    run.font.size = Pt(24)

out = "/home/azureuser/.openclaw/workspace/pptx-editor/test_report.pptx"
prs.save(out)
print(f"✅ 测试 PPTX 生成: {out}")
