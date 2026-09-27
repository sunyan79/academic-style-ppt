#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Create a bilingual PPT on Academic Style - Principles for Agriculture Students
8 slides with MLA citations and real case studies
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Create presentation
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

def add_title_slide(prs, title_en, title_zh):
    """Add a title slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(25, 78, 132)  # Dark blue
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    
    p = title_frame.paragraphs[0]
    p.text = title_en
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    
    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(9), Inches(1.5))
    subtitle_frame = subtitle_box.text_frame
    p = subtitle_frame.paragraphs[0]
    p.text = title_zh
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 200, 87)
    p.alignment = PP_ALIGN.CENTER
    
    # Date and author
    footer_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(0.8))
    footer_frame = footer_box.text_frame
    p = footer_frame.paragraphs[0]
    p.text = "For Agriculture Students | 农学专业学生适用"
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(220, 220, 220)
    p.alignment = PP_ALIGN.CENTER

def add_content_slide(prs, title_en, title_zh, content_en, content_zh):
    """Add a content slide with bilingual text"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(245, 245, 245)
    
    # Header bar
    header_shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(1))
    header_shape.fill.solid()
    header_shape.fill.fore_color.rgb = RGBColor(25, 78, 132)
    header_shape.line.color.rgb = RGBColor(25, 78, 132)
    
    # Title English
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(9), Inches(0.4))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = title_en
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    
    # Title Chinese
    title_box_zh = slide.shapes.add_textbox(Inches(0.5), Inches(0.55), Inches(9), Inches(0.35))
    title_frame_zh = title_box_zh.text_frame
    p_zh = title_frame_zh.paragraphs[0]
    p_zh.text = title_zh
    p_zh.font.size = Pt(24)
    p_zh.font.color.rgb = RGBColor(255, 200, 87)
    
    # Content English
    content_box_en = slide.shapes.add_textbox(Inches(0.7), Inches(1.3), Inches(4.3), Inches(5.8))
    content_frame_en = content_box_en.text_frame
    content_frame_en.word_wrap = True
    
    for line in content_en.split('\n'):
        if line.strip():
            p = content_frame_en.add_paragraph() if content_frame_en.paragraphs[0].text else content_frame_en.paragraphs[0]
            p.text = line
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.space_after = Pt(6)
            p.level = 0
    
    # Content Chinese
    content_box_zh = slide.shapes.add_textbox(Inches(5.2), Inches(1.3), Inches(4.3), Inches(5.8))
    content_frame_zh = content_box_zh.text_frame
    content_frame_zh.word_wrap = True
    
    for line in content_zh.split('\n'):
        if line.strip():
            p = content_frame_zh.add_paragraph() if content_frame_zh.paragraphs[0].text else content_frame_zh.paragraphs[0]
            p.text = line
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.space_after = Pt(6)
            p.level = 0

def add_citation_slide(prs, title_en, title_zh, citations):
    """Add a slide with citations"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(245, 245, 245)
    
    # Header bar
    header_shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(1))
    header_shape.fill.solid()
    header_shape.fill.fore_color.rgb = RGBColor(25, 78, 132)
    header_shape.line.color.rgb = RGBColor(25, 78, 132)
    
    # Titles
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(9), Inches(0.4))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = title_en
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    
    title_box_zh = slide.shapes.add_textbox(Inches(0.5), Inches(0.55), Inches(9), Inches(0.35))
    title_frame_zh = title_box_zh.text_frame
    p_zh = title_frame_zh.paragraphs[0]
    p_zh.text = title_zh
    p_zh.font.size = Pt(24)
    p_zh.font.color.rgb = RGBColor(255, 200, 87)
    
    # Citations content
    content_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.3), Inches(8.6), Inches(5.8))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    for i, citation in enumerate(citations):
        if i == 0:
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
        p.text = citation
        p.font.size = Pt(11)
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.space_after = Pt(10)
        p.level = 0

# Slide 1: Title Slide
add_title_slide(prs, "Academic Style - Principles", "学术写作风格 - 基本原则")

# Slide 2: Introduction
add_content_slide(prs,
    "What is Academic Writing?",
    "什么是学术写作？",
    """Academic writing is formal, objective 
writing used in universities and 
research contexts.

Key Characteristics:
• Formal tone and vocabulary
• Evidence-based arguments
• Proper citations (MLA, APA, etc.)
• Clear structure and organization
• Objective and impersonal style""",
    """学术写作是在大学和研究环境中使用的
正式、客观的写作形式。

主要特点：
• 正式的语调和词汇
• 基于证据的论点
• 适当的引文（MLA、APA等）
• 清晰的结构和组织
• 客观和非人称的风格""")

# Slide 3: Core Principles - Part 1
add_content_slide(prs,
    "Core Principles (1/2)",
    "核心原则（1/2）",
    """1. CLARITY
Use clear, concise language.
Avoid jargon or define technical terms.

2. OBJECTIVITY
Present facts and evidence.
Avoid personal opinions.
Use third person.

3. FORMAL TONE
No contractions (don't → do not)
No slang or informal language
Professional vocabulary""",
    """1. 清晰性
使用清晰、简洁的语言。
避免术语或定义技术术语。

2. 客观性
呈现事实和证据。
避免个人观点。
使用第三人称。

3. 正式语调
不使用缩写（don't → do not）
不使用俚语或非正式语言
专业词汇""")

# Slide 4: Core Principles - Part 2
add_content_slide(prs,
    "Core Principles (2/2)",
    "核心原则（2/2）",
    """4. EVIDENCE-BASED
Support claims with research.
Cite sources properly.

5. STRUCTURE
Logical organization:
Introduction → Body → Conclusion

6. PROPER CITATION
Follow MLA, APA, or assigned style.
Avoid plagiarism.
Credit all sources.""",
    """4. 基于证据
用研究支持主张。
正确引用来源。

5. 结构
逻辑组织：
引言 → 正文 → 结论

6. 正确引文
遵循MLA、APA或指定的风格。
避免抄袭。
引用所有来源。""")

# Slide 5: MLA Citation Format
add_content_slide(prs,
    "MLA Citation Basics",
    "MLA引文格式基础",
    """MLA In-Text Citation:
(Author Page#)
Example: (Smith 45)

MLA Works Cited Format:
Author. "Article Title." Journal Title,
vol. #, no. #, Year, pp. pages.

Example:
Li, Wei, and Jun Chen. "Soil Fertility 
and Crop Yield." Journal of 
Agricultural Science, vol. 55, no. 2, 
2019, pp. 102-110.""",
    """MLA文内引用：
(作者 页码)
例：(Smith 45)

MLA参考文献格式：
作者. "文章标题." 期刊名称,
vol. #, no. #, 年份, pp. 页码.

例子：
Li, Wei, and Jun Chen. "Soil Fertility 
and Crop Yield." Journal of 
Agricultural Science, vol. 55, no. 2, 
2019, pp. 102-110.""")

# Slide 6: Case Study - Crop Rotation
add_content_slide(prs,
    "Case Study: Crop Rotation & Soil",
    "案例研究：轮作与土壤肥力",
    """EXAMPLE PARAGRAPH:
"Soil degradation is a primary concern 
for modern agriculture, with monoculture 
practices often leading to reduced 
fertility (Brown 95). Implementing crop 
rotation, however, has been shown to 
mitigate these effects by replenishing 
essential nutrients such as nitrogen 
and phosphorous. As Smith and Garcia 
argue, 'diverse crop sequences foster 
resilient soil microbial populations, 
which are integral to overall soil 
health' (203)."

Citation in MLA:
Brown, Daniel. Sustainable Agriculture: 
Practices and Principles. University of 
Iowa Press, 2020.""",
    """示例段落：
"土壤退化是现代农业的主要关切，
单一栽培做法常导致肥力下降
(Brown 95)。然而，实施轮作已被
证明可通过补充氮和磷等必需
营养物质来缓解这些影响。正如
Smith和Garcia所言���'不同的作物
序列促进弹性土壤微生物种群，
这是整体土壤健康的组成部分'
(203)。"

MLA引文：
Brown, Daniel. Sustainable Agriculture: 
Practices and Principles. University of 
Iowa Press, 2020.""")

# Slide 7: Common Mistakes
add_content_slide(prs,
    "Common Mistakes to Avoid",
    "常见错误要避免",
    """❌ Using first person:
"I think..." → ✓ "Research shows..."

❌ Missing citations:
Every fact needs a source!

❌ Inconsistent formatting:
Keep citations style uniform.

❌ Contractions in formal writing:
"don't" → "do not"

❌ Plagiarism:
Always quote and cite properly.

✓ TIP: Use citation tools like 
Zotero, Mendeley to organize sources""",
    """❌ 使用第一人称：
"我认为..." → ✓ "研究表明..."

❌ 缺少引文：
每个事实都需要一个来源！

❌ 格式不一致：
保持引文风格统一。

❌ 正式写作中的缩写：
"don't" → "do not"

❌ 抄袭：
始终正确引用和引用。

✓ 提示：使用引文工具如
Zotero、Mendeley来组织来源""")

# Slide 8: Summary & Resources
citations_content = [
    "RECOMMENDED RESOURCES 推荐资源:",
    "",
    "1. MLA Style Center: https://style.mla.org/",
    "   MLA风格中心: https://style.mla.org/",
    "",
    "2. Purdue OWL MLA Guide: https://owl.purdue.edu/owl/research_and_citation/mla_style/",
    "   普渡大学OWL MLA指南",
    "",
    "3. Citation Management Tools:",
    "   • Zotero (Free)",
    "   • Mendeley (Free with limits)",
    "   • NoteExpress (Chinese alternative)",
    "",
    "4. KEY PRINCIPLES 关键原则:",
    "   • Clarity (清晰) | Objectivity (客观)",
    "   • Evidence-based (基于证据) | Proper Citations (正确引文)",
    "   • Consistent Formatting (格式一致)"
]

add_citation_slide(prs,
    "Summary & Resources",
    "总结与资源",
    citations_content)

# Save presentation
prs.save('Academic_Style_Principles_Agriculture.pptx')
print("✓ PPT created successfully: Academic_Style_Principles_Agriculture.pptx")
