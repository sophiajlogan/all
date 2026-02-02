#!/usr/bin/env python3
"""
Create Enterprise Analytics and AI Team Introduction PowerPoint
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor as RgbColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Create presentation with widescreen dimensions
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Color scheme - professional blues and greens (renewable energy feel)
PRIMARY_BLUE = RgbColor(0, 82, 147)  # Deep blue
ACCENT_GREEN = RgbColor(0, 150, 136)  # Teal green
LIGHT_BLUE = RgbColor(66, 133, 244)  # Lighter blue
DARK_GRAY = RgbColor(60, 60, 60)
WHITE = RgbColor(255, 255, 255)
LIGHT_BG = RgbColor(245, 248, 250)

def add_title_shape(slide, text, top, font_size=44, color=PRIMARY_BLUE):
    """Add a title text box"""
    shape = slide.shapes.add_textbox(Inches(0.5), top, Inches(12.333), Inches(1))
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = True
    p.font.color.rgb = color
    p.alignment = PP_ALIGN.LEFT
    return shape

def add_body_text(slide, text, left, top, width, height, font_size=18, color=DARK_GRAY, bold=False):
    """Add body text box"""
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.alignment = PP_ALIGN.LEFT
    return shape

def add_bullet_points(slide, bullets, left, top, width, height, font_size=16):
    """Add bullet point text"""
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.word_wrap = True

    for i, bullet in enumerate(bullets):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = "• " + bullet
        p.font.size = Pt(font_size)
        p.font.color.rgb = DARK_GRAY
        p.space_after = Pt(8)
    return shape

# ============================================
# SLIDE 1: Title Slide
# ============================================
slide_layout = prs.slide_layouts[6]  # Blank
slide1 = prs.slides.add_slide(slide_layout)

# Background shape
bg = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
bg.fill.solid()
bg.fill.fore_color.rgb = PRIMARY_BLUE
bg.line.fill.background()

# Accent bar
accent = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(5.2), Inches(13.333), Inches(0.15))
accent.fill.solid()
accent.fill.fore_color.rgb = ACCENT_GREEN
accent.line.fill.background()

# Title
title_box = slide1.shapes.add_textbox(Inches(0.75), Inches(2.5), Inches(11.833), Inches(1.5))
tf = title_box.text_frame
p = tf.paragraphs[0]
p.text = "Enterprise Analytics and AI"
p.font.size = Pt(54)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.LEFT

# Subtitle
sub_box = slide1.shapes.add_textbox(Inches(0.75), Inches(4.2), Inches(11.833), Inches(1))
tf = sub_box.text_frame
p = tf.paragraphs[0]
p.text = "Your Partners in Data-Driven Development"
p.font.size = Pt(28)
p.font.color.rgb = RgbColor(200, 220, 240)
p.alignment = PP_ALIGN.LEFT

# Bottom tagline
tag_box = slide1.shapes.add_textbox(Inches(0.75), Inches(5.8), Inches(11.833), Inches(1))
tf = tag_box.text_frame
p = tf.paragraphs[0]
p.text = "Helping teams turn data into insights, and insights into action"
p.font.size = Pt(18)
p.font.italic = True
p.font.color.rgb = RgbColor(180, 200, 220)
p.alignment = PP_ALIGN.LEFT

# ============================================
# SLIDE 2: Who We Are - The Metaphor
# ============================================
slide2 = prs.slides.add_slide(slide_layout)

add_title_shape(slide2, "Who We Are", Inches(0.4))

# Subtitle/hook
hook_box = slide2.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(12), Inches(0.8))
tf = hook_box.text_frame
p = tf.paragraphs[0]
p.text = "Think of us as your in-house technology translators"
p.font.size = Pt(24)
p.font.italic = True
p.font.color.rgb = ACCENT_GREEN

# Main metaphor box
metaphor_bg = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(2), Inches(12.333), Inches(2.2))
metaphor_bg.fill.solid()
metaphor_bg.fill.fore_color.rgb = RgbColor(240, 248, 255)
metaphor_bg.line.color.rgb = LIGHT_BLUE

metaphor_text = slide2.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(11.8), Inches(2))
tf = metaphor_text.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Imagine having a neighbor who could build you a custom website..."
p.font.size = Pt(20)
p.font.color.rgb = DARK_GRAY
p.space_after = Pt(12)

p2 = tf.add_paragraph()
p2.text = "Now imagine that neighbor also knows how to build the systems that help Netflix figure out which movies to recommend to you, or how Amazon predicts what you might want to buy next."
p2.font.size = Pt(20)
p2.font.color.rgb = DARK_GRAY
p2.space_after = Pt(12)

p3 = tf.add_paragraph()
p3.text = "That's us — technologists with computer science and data science backgrounds who specialize in finding patterns in data and building tools that make complex information useful."
p3.font.size = Pt(20)
p3.font.bold = True
p3.font.color.rgb = PRIMARY_BLUE

# Three column layout for key points
col_width = Inches(3.8)
col_start = Inches(0.6)
col_top = Inches(4.5)

# Column 1
box1 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_start, col_top, col_width, Inches(2.5))
box1.fill.solid()
box1.fill.fore_color.rgb = PRIMARY_BLUE
box1.line.fill.background()

text1 = slide2.shapes.add_textbox(col_start + Inches(0.2), col_top + Inches(0.2), col_width - Inches(0.4), Inches(2.3))
tf = text1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "🔧 We Build"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = WHITE
p.space_after = Pt(8)
p2 = tf.add_paragraph()
p2.text = "Databases, dashboards, automated reports, forecasting models, and AI-powered tools"
p2.font.size = Pt(16)
p2.font.color.rgb = WHITE

# Column 2
box2 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_start + Inches(4.2), col_top, col_width, Inches(2.5))
box2.fill.solid()
box2.fill.fore_color.rgb = ACCENT_GREEN
box2.line.fill.background()

text2 = slide2.shapes.add_textbox(col_start + Inches(4.4), col_top + Inches(0.2), col_width - Inches(0.4), Inches(2.3))
tf = text2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "🔍 We Analyze"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = WHITE
p.space_after = Pt(8)
p2 = tf.add_paragraph()
p2.text = "Find trends, spot patterns, predict outcomes, and turn messy data into clear answers"
p2.font.size = Pt(16)
p2.font.color.rgb = WHITE

# Column 3
box3 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_start + Inches(8.4), col_top, col_width, Inches(2.5))
box3.fill.solid()
box3.fill.fore_color.rgb = LIGHT_BLUE
box3.line.fill.background()

text3 = slide2.shapes.add_textbox(col_start + Inches(8.6), col_top + Inches(0.2), col_width - Inches(0.4), Inches(2.3))
tf = text3.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "🤝 We Partner"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = WHITE
p.space_after = Pt(8)
p2 = tf.add_paragraph()
p2.text = "Work alongside you to understand your challenges and create solutions that actually work"
p2.font.size = Pt(16)
p2.font.color.rgb = WHITE

# ============================================
# SLIDE 3: Hub and Spoke Diagram
# ============================================
slide3 = prs.slides.add_slide(slide_layout)

add_title_shape(slide3, "How We Work With You", Inches(0.3))

# Subtitle explaining the model
sub = slide3.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(12), Inches(0.6))
tf = sub.text_frame
p = tf.paragraphs[0]
p.text = "We're a shared resource that amplifies your team's capabilities — not a replacement"
p.font.size = Pt(18)
p.font.color.rgb = DARK_GRAY

# Center hub
hub_x = Inches(6.666) - Inches(1.1)
hub_y = Inches(4) - Inches(1.1)
hub = slide3.shapes.add_shape(MSO_SHAPE.OVAL, hub_x, hub_y, Inches(2.2), Inches(2.2))
hub.fill.solid()
hub.fill.fore_color.rgb = PRIMARY_BLUE
hub.line.fill.background()

hub_text = slide3.shapes.add_textbox(hub_x + Inches(0.15), hub_y + Inches(0.7), Inches(1.9), Inches(0.8))
tf = hub_text.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Enterprise Analytics & AI"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

# Spoke teams - positioned around the hub
spokes = [
    ("Land\nAcquisition", Inches(2.5), Inches(2)),
    ("Permitting", Inches(10), Inches(2)),
    ("Engineering", Inches(10.5), Inches(5)),
    ("Finance", Inches(2), Inches(5.2)),
    ("Legal", Inches(5.5), Inches(6.3)),
    ("Operations", Inches(7.5), Inches(6.3)),
]

spoke_size = Inches(1.5)

for name, x, y in spokes:
    # Draw connector line (behind the spoke)
    # Calculate center points
    spoke_center_x = x + spoke_size/2
    spoke_center_y = y + spoke_size/2
    hub_center_x = hub_x + Inches(1.1)
    hub_center_y = hub_y + Inches(1.1)

    # Add connector as a line
    connector = slide3.shapes.add_connector(
        1,  # Straight connector
        spoke_center_x, spoke_center_y,
        hub_center_x, hub_center_y
    )
    connector.line.color.rgb = RgbColor(180, 200, 220)
    connector.line.width = Pt(2)

    # Spoke circle
    spoke = slide3.shapes.add_shape(MSO_SHAPE.OVAL, x, y, spoke_size, spoke_size)
    spoke.fill.solid()
    spoke.fill.fore_color.rgb = ACCENT_GREEN
    spoke.line.fill.background()

    # Spoke text
    spoke_text = slide3.shapes.add_textbox(x + Inches(0.1), y + Inches(0.45), spoke_size - Inches(0.2), Inches(0.6))
    tf = spoke_text.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = name
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

# Key message boxes at bottom
msg_box1 = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.6), Inches(3.8), Inches(0.9))
msg_box1.fill.solid()
msg_box1.fill.fore_color.rgb = RgbColor(232, 245, 233)
msg_box1.line.color.rgb = ACCENT_GREEN

msg1 = slide3.shapes.add_textbox(Inches(0.7), Inches(1.75), Inches(3.4), Inches(0.7))
tf = msg1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "✓ You keep your domain expertise"
p.font.size = Pt(14)
p.font.color.rgb = DARK_GRAY

msg_box2 = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.5), Inches(1.6), Inches(3.8), Inches(0.9))
msg_box2.fill.solid()
msg_box2.fill.fore_color.rgb = RgbColor(232, 245, 233)
msg_box2.line.color.rgb = ACCENT_GREEN

msg2 = slide3.shapes.add_textbox(Inches(4.7), Inches(1.75), Inches(3.4), Inches(0.7))
tf = msg2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "✓ We bring technology skills"
p.font.size = Pt(14)
p.font.color.rgb = DARK_GRAY

msg_box3 = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.5), Inches(1.6), Inches(4.3), Inches(0.9))
msg_box3.fill.solid()
msg_box3.fill.fore_color.rgb = RgbColor(232, 245, 233)
msg_box3.line.color.rgb = ACCENT_GREEN

msg3 = slide3.shapes.add_textbox(Inches(8.7), Inches(1.75), Inches(3.9), Inches(0.7))
tf = msg3.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "✓ Together we build cross-team tools"
p.font.size = Pt(14)
p.font.color.rgb = DARK_GRAY

# ============================================
# SLIDE 4: Internal Consultants Model
# ============================================
slide4 = prs.slides.add_slide(slide_layout)

add_title_shape(slide4, "We're Like Internal Consultants", Inches(0.3))

sub = slide4.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(12), Inches(0.6))
tf = sub.text_frame
p = tf.paragraphs[0]
p.text = "You're the expert in your work. We're experts in building technology to support it."
p.font.size = Pt(20)
p.font.italic = True
p.font.color.rgb = ACCENT_GREEN

# Left side - What you bring
left_header = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.9), Inches(5.5), Inches(0.7))
left_header.fill.solid()
left_header.fill.fore_color.rgb = ACCENT_GREEN
left_header.line.fill.background()

left_title = slide4.shapes.add_textbox(Inches(0.7), Inches(2), Inches(5.1), Inches(0.5))
tf = left_title.text_frame
p = tf.paragraphs[0]
p.text = "What You Bring"
p.font.size = Pt(22)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

left_box = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(2.6), Inches(5.5), Inches(2.8))
left_box.fill.solid()
left_box.fill.fore_color.rgb = RgbColor(240, 248, 245)
left_box.line.color.rgb = ACCENT_GREEN

left_bullets = [
    "Deep knowledge of land acquisition, permitting, or your specialty",
    "Understanding of what questions need answers",
    "Relationships with landowners, agencies, stakeholders",
    "Experience navigating complex project requirements"
]
add_bullet_points(slide4, left_bullets, Inches(0.7), Inches(2.8), Inches(5.1), Inches(2.5), font_size=15)

# Right side - What we bring
right_header = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.5), Inches(1.9), Inches(6.3), Inches(0.7))
right_header.fill.solid()
right_header.fill.fore_color.rgb = PRIMARY_BLUE
right_header.line.fill.background()

right_title = slide4.shapes.add_textbox(Inches(6.7), Inches(2), Inches(5.9), Inches(0.5))
tf = right_title.text_frame
p = tf.paragraphs[0]
p.text = "What We Bring"
p.font.size = Pt(22)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

right_box = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.5), Inches(2.6), Inches(6.3), Inches(2.8))
right_box.fill.solid()
right_box.fill.fore_color.rgb = RgbColor(240, 245, 250)
right_box.line.color.rgb = PRIMARY_BLUE

right_bullets = [
    "Ability to connect scattered data sources into one view",
    "Skills to automate repetitive data tasks",
    "Experience building predictive models and AI tools",
    "Knowledge of best practices from across the industry"
]
add_bullet_points(slide4, right_bullets, Inches(6.7), Inches(2.8), Inches(5.9), Inches(2.5), font_size=15)

# Bottom - Together box
together_header = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.6))
together_header.fill.solid()
together_header.fill.fore_color.rgb = LIGHT_BLUE
together_header.line.fill.background()

together_title = slide4.shapes.add_textbox(Inches(0.7), Inches(5.68), Inches(11.9), Inches(0.5))
tf = together_title.text_frame
p = tf.paragraphs[0]
p.text = "Together We Create"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

together_box = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(6.2), Inches(12.3), Inches(1.1))
together_box.fill.solid()
together_box.fill.fore_color.rgb = RgbColor(232, 240, 254)
together_box.line.color.rgb = LIGHT_BLUE

together_text = slide4.shapes.add_textbox(Inches(0.7), Inches(6.35), Inches(11.9), Inches(0.9))
tf = together_text.text_frame
p = tf.paragraphs[0]
p.text = "Solutions that actually solve your real problems — because we build them together based on your expertise"
p.font.size = Pt(18)
p.font.color.rgb = DARK_GRAY
p.alignment = PP_ALIGN.CENTER

# ============================================
# SLIDE 5: What We Can Build - Relevant Examples
# ============================================
slide5 = prs.slides.add_slide(slide_layout)

add_title_shape(slide5, "What Can We Build Together?", Inches(0.3))

sub = slide5.shapes.add_textbox(Inches(0.5), Inches(1), Inches(12), Inches(0.6))
tf = sub.text_frame
p = tf.paragraphs[0]
p.text = "Examples relevant to getting projects from land acquisition through permitting"
p.font.size = Pt(16)
p.font.color.rgb = DARK_GRAY

# Example cards - 2x3 grid
examples = [
    ("📊", "Unified Project Dashboard",
     "One place to see every project's status from initial land contact through permit approval"),
    ("🗄️", "Connected Databases",
     "Link your spreadsheets, documents, and systems so data flows automatically instead of manual updates"),
    ("📈", "Trend Analysis",
     "See patterns like which counties approve permits faster, or what landowner characteristics predict success"),
    ("🔮", "Timeline Predictions",
     "Forecast realistic permit approval dates based on historical data and current conditions"),
    ("🤖", "AI Document Tools",
     "Auto-generate permit applications, landowner letters, or reports from your project data"),
    ("🗺️", "Visual Mapping",
     "Interactive maps showing project pipeline, land status, and key metrics at a glance"),
]

card_width = Inches(4)
card_height = Inches(2.2)
start_x = Inches(0.5)
start_y = Inches(1.6)
gap_x = Inches(0.2)
gap_y = Inches(0.2)

for i, (icon, title, desc) in enumerate(examples):
    row = i // 3
    col = i % 3
    x = start_x + col * (card_width + gap_x)
    y = start_y + row * (card_height + gap_y)

    # Card background
    card = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_width, card_height)
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = RgbColor(200, 210, 220)
    card.line.width = Pt(1)

    # Colored top bar
    colors = [PRIMARY_BLUE, ACCENT_GREEN, LIGHT_BLUE, RgbColor(156, 39, 176), RgbColor(255, 152, 0), RgbColor(76, 175, 80)]
    top_bar = slide5.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, card_width, Inches(0.15))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = colors[i]
    top_bar.line.fill.background()

    # Icon and title
    title_box = slide5.shapes.add_textbox(x + Inches(0.15), y + Inches(0.25), card_width - Inches(0.3), Inches(0.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = f"{icon} {title}"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = DARK_GRAY

    # Description
    desc_box = slide5.shapes.add_textbox(x + Inches(0.15), y + Inches(0.75), card_width - Inches(0.3), Inches(1.4))
    tf = desc_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = desc
    p.font.size = Pt(13)
    p.font.color.rgb = RgbColor(100, 100, 100)

# Bottom note
note = slide5.shapes.add_textbox(Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.6))
tf = note.text_frame
p = tf.paragraphs[0]
p.text = "These are just examples — we tailor solutions to your specific challenges"
p.font.size = Pt(16)
p.font.italic = True
p.font.color.rgb = ACCENT_GREEN
p.alignment = PP_ALIGN.CENTER

# ============================================
# SLIDE 6: How to Work With Us
# ============================================
slide6 = prs.slides.add_slide(slide_layout)

add_title_shape(slide6, "How to Work With Us", Inches(0.3))

# Process steps - horizontal flow
steps = [
    ("1", "Share Your Challenge", "Tell us about a problem you're facing or an idea you have"),
    ("2", "We Explore Together", "We ask questions to understand your world and what success looks like"),
    ("3", "Design a Solution", "We propose approaches and you help us refine them"),
    ("4", "Build & Iterate", "We create, you test, we improve together"),
]

step_width = Inches(2.8)
step_height = Inches(3.5)
start_x = Inches(0.6)
arrow_width = Inches(0.4)

for i, (num, title, desc) in enumerate(steps):
    x = start_x + i * (step_width + arrow_width)
    y = Inches(1.5)

    # Number circle
    circle = slide6.shapes.add_shape(MSO_SHAPE.OVAL, x + step_width/2 - Inches(0.4), y, Inches(0.8), Inches(0.8))
    circle.fill.solid()
    circle.fill.fore_color.rgb = PRIMARY_BLUE if i % 2 == 0 else ACCENT_GREEN
    circle.line.fill.background()

    num_text = slide6.shapes.add_textbox(x + step_width/2 - Inches(0.35), y + Inches(0.15), Inches(0.7), Inches(0.5))
    tf = num_text.text_frame
    p = tf.paragraphs[0]
    p.text = num
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

    # Title
    title_box = slide6.shapes.add_textbox(x, y + Inches(1), step_width, Inches(0.6))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = DARK_GRAY
    p.alignment = PP_ALIGN.CENTER

    # Description
    desc_box = slide6.shapes.add_textbox(x, y + Inches(1.6), step_width, Inches(1.5))
    tf = desc_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = desc
    p.font.size = Pt(14)
    p.font.color.rgb = RgbColor(100, 100, 100)
    p.alignment = PP_ALIGN.CENTER

    # Arrow between steps
    if i < len(steps) - 1:
        arrow = slide6.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x + step_width + Inches(0.05), y + Inches(0.2), arrow_width - Inches(0.1), Inches(0.4))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = RgbColor(200, 210, 220)
        arrow.line.fill.background()

# Bottom message box
msg_bg = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.3), Inches(12.3), Inches(1.8))
msg_bg.fill.solid()
msg_bg.fill.fore_color.rgb = RgbColor(240, 248, 255)
msg_bg.line.color.rgb = PRIMARY_BLUE

msg_title = slide6.shapes.add_textbox(Inches(0.7), Inches(5.5), Inches(11.9), Inches(0.5))
tf = msg_title.text_frame
p = tf.paragraphs[0]
p.text = "No request is too small or too ambitious"
p.font.size = Pt(22)
p.font.bold = True
p.font.color.rgb = PRIMARY_BLUE
p.alignment = PP_ALIGN.CENTER

msg_body = slide6.shapes.add_textbox(Inches(0.7), Inches(6.1), Inches(11.9), Inches(0.9))
tf = msg_body.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Whether it's \"Can you help me understand this spreadsheet?\" or \"We need a system to track 500 projects\" — we're here to help. The best solutions often start with a simple conversation."
p.font.size = Pt(16)
p.font.color.rgb = DARK_GRAY
p.alignment = PP_ALIGN.CENTER

# Save the presentation
prs.save('/home/user/all/Enterprise_Analytics_AI_Team.pptx')
print("Presentation saved to: /home/user/all/Enterprise_Analytics_AI_Team.pptx")
