#!/usr/bin/env python3
"""
Create Enterprise Analytics and AI Team Introduction PowerPoint
Polished version with service area details
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

# Color scheme - refined professional palette
NAVY = RgbColor(23, 42, 69)         # Deep navy for headers
TEAL = RgbColor(0, 128, 128)        # Teal accent
SKY = RgbColor(70, 130, 180)        # Steel blue
SAGE = RgbColor(85, 139, 110)       # Muted green (renewable energy)
CORAL = RgbColor(205, 92, 92)       # Warm accent
SLATE = RgbColor(71, 85, 105)       # Dark gray for text
LIGHT_GRAY = RgbColor(248, 250, 252)
MED_GRAY = RgbColor(100, 116, 139)
WHITE = RgbColor(255, 255, 255)

def add_slide_footer(slide, text="Enterprise Analytics and AI"):
    """Add subtle footer to slide"""
    footer = slide.shapes.add_textbox(Inches(0.5), Inches(7.1), Inches(4), Inches(0.3))
    tf = footer.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(9)
    p.font.color.rgb = MED_GRAY
    p.font.italic = True

# ============================================
# SLIDE 1: Title Slide - Clean & Professional
# ============================================
slide1 = prs.slides.add_slide(prs.slide_layouts[6])

# Gradient-like background using layered shapes
bg = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
bg.fill.solid()
bg.fill.fore_color.rgb = NAVY
bg.line.fill.background()

# Subtle geometric accent
accent1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.5), Inches(13.333), Inches(0.08))
accent1.fill.solid()
accent1.fill.fore_color.rgb = TEAL
accent1.line.fill.background()

accent2 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.65), Inches(8), Inches(0.04))
accent2.fill.solid()
accent2.fill.fore_color.rgb = SAGE
accent2.line.fill.background()

# Main title
title_box = slide1.shapes.add_textbox(Inches(0.8), Inches(2.6), Inches(11.733), Inches(1.2))
tf = title_box.text_frame
p = tf.paragraphs[0]
p.text = "Enterprise Analytics and AI"
p.font.size = Pt(52)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.LEFT

# Subtitle
sub_box = slide1.shapes.add_textbox(Inches(0.8), Inches(4.0), Inches(11.733), Inches(0.8))
tf = sub_box.text_frame
p = tf.paragraphs[0]
p.text = "Partnering with teams to turn data into decisions"
p.font.size = Pt(24)
p.font.color.rgb = RgbColor(176, 196, 222)
p.alignment = PP_ALIGN.LEFT

# Bottom tagline
tag_box = slide1.shapes.add_textbox(Inches(0.8), Inches(6.2), Inches(11.733), Inches(0.6))
tf = tag_box.text_frame
p = tf.paragraphs[0]
p.text = "From land acquisition through permitting — we help you see the full picture"
p.font.size = Pt(16)
p.font.italic = True
p.font.color.rgb = RgbColor(140, 160, 180)
p.alignment = PP_ALIGN.LEFT

# ============================================
# SLIDE 2: Who We Are - Plain Language
# ============================================
slide2 = prs.slides.add_slide(prs.slide_layouts[6])

# Light background
bg2 = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
bg2.fill.solid()
bg2.fill.fore_color.rgb = WHITE
bg2.line.fill.background()

# Header bar
header = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.3))
header.fill.solid()
header.fill.fore_color.rgb = NAVY
header.line.fill.background()

title = slide2.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.8))
tf = title.text_frame
p = tf.paragraphs[0]
p.text = "Who We Are"
p.font.size = Pt(36)
p.font.bold = True
p.font.color.rgb = WHITE

# Main metaphor section - shortened
metaphor_box = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.7), Inches(12.133), Inches(1.4))
metaphor_box.fill.solid()
metaphor_box.fill.fore_color.rgb = RgbColor(240, 248, 255)
metaphor_box.line.color.rgb = SKY
metaphor_box.line.width = Pt(1.5)

metaphor_text = slide2.shapes.add_textbox(Inches(0.9), Inches(1.9), Inches(11.5), Inches(1.1))
tf = metaphor_text.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "You know how Netflix figures out which shows you'll like, or how your phone identifies a song in seconds? We build those kinds of systems — but for your business problems."
p.font.size = Pt(20)
p.font.color.rgb = SLATE

# Three key descriptors
col_width = Inches(3.9)
col_gap = Inches(0.25)
col_y = Inches(3.4)

descriptors = [
    (TEAL, "We Build Tools", "From simple dashboards to smart systems that learn and improve over time"),
    (SKY, "We Analyze Data", "Finding the patterns and trends hidden in your spreadsheets and databases"),
    (SAGE, "We Solve Problems", "Working alongside you to tackle challenges with the right technology")
]

for i, (color, title_text, desc) in enumerate(descriptors):
    x = Inches(0.6) + i * (col_width + col_gap)

    # Card
    card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, col_y, col_width, Inches(2.1))
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = color
    card.line.width = Pt(2)

    # Color bar at top
    bar = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, col_y, col_width, Inches(0.12))
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()

    # Title
    t = slide2.shapes.add_textbox(x + Inches(0.2), col_y + Inches(0.3), col_width - Inches(0.4), Inches(0.5))
    tf = t.text_frame
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = color

    # Description
    d = slide2.shapes.add_textbox(x + Inches(0.2), col_y + Inches(0.85), col_width - Inches(0.4), Inches(1.1))
    tf = d.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = desc
    p.font.size = Pt(14)
    p.font.color.rgb = SLATE

# Bottom note
note = slide2.shapes.add_textbox(Inches(0.6), Inches(5.8), Inches(12), Inches(0.6))
tf = note.text_frame
p = tf.paragraphs[0]
p.text = "Computer science + data science backgrounds → Real-world solutions for your team"
p.font.size = Pt(15)
p.font.italic = True
p.font.color.rgb = MED_GRAY
p.alignment = PP_ALIGN.CENTER

add_slide_footer(slide2)

# ============================================
# SLIDE 3: Hub and Spoke - How We Support Teams
# ============================================
slide3 = prs.slides.add_slide(prs.slide_layouts[6])

bg3 = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
bg3.fill.solid()
bg3.fill.fore_color.rgb = WHITE
bg3.line.fill.background()

header3 = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.3))
header3.fill.solid()
header3.fill.fore_color.rgb = NAVY
header3.line.fill.background()

title3 = slide3.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.8))
tf = title3.text_frame
p = tf.paragraphs[0]
p.text = "How We Work With You"
p.font.size = Pt(36)
p.font.bold = True
p.font.color.rgb = WHITE

# Subtitle
sub3 = slide3.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.5), Inches(0.5))
tf = sub3.text_frame
p = tf.paragraphs[0]
p.text = "A shared resource that amplifies your capabilities — we don't replace your expertise, we enhance it"
p.font.size = Pt(16)
p.font.color.rgb = MED_GRAY

# Hub and Spoke Diagram
hub_center_x = Inches(6.666)
hub_center_y = Inches(4.3)
hub_radius = Inches(0.9)

# Center hub
hub = slide3.shapes.add_shape(MSO_SHAPE.OVAL,
    hub_center_x - hub_radius,
    hub_center_y - hub_radius,
    hub_radius * 2, hub_radius * 2)
hub.fill.solid()
hub.fill.fore_color.rgb = NAVY
hub.line.fill.background()

hub_text = slide3.shapes.add_textbox(hub_center_x - Inches(0.75), hub_center_y - Inches(0.35), Inches(1.5), Inches(0.7))
tf = hub_text.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Enterprise\nAnalytics & AI"
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

# Spoke positions (around the hub)
import math
spokes = [
    "Land\nAcquisition",
    "Permitting",
    "Engineering",
    "Finance",
    "Legal",
    "Operations"
]

spoke_radius = Inches(2.3)
spoke_size = Inches(1.2)

for i, name in enumerate(spokes):
    angle = (i * 60 - 90) * math.pi / 180  # Start from top, go clockwise
    x = hub_center_x + spoke_radius * math.cos(angle) - spoke_size/2
    y = hub_center_y + spoke_radius * math.sin(angle) - spoke_size/2

    # Connector line
    spoke_cx = x + spoke_size/2
    spoke_cy = y + spoke_size/2
    connector = slide3.shapes.add_connector(1, spoke_cx, spoke_cy, hub_center_x, hub_center_y)
    connector.line.color.rgb = RgbColor(200, 215, 230)
    connector.line.width = Pt(2)

    # Spoke circle
    spoke = slide3.shapes.add_shape(MSO_SHAPE.OVAL, x, y, spoke_size, spoke_size)
    spoke.fill.solid()
    spoke.fill.fore_color.rgb = TEAL
    spoke.line.fill.background()

    # Spoke text
    st = slide3.shapes.add_textbox(x + Inches(0.1), y + spoke_size/2 - Inches(0.25), spoke_size - Inches(0.2), Inches(0.5))
    tf = st.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = name
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

# Key points - left and right of diagram
left_points = [
    "✓ You remain the expert in your domain",
    "✓ We add technology capabilities",
    "✓ Solutions built for your actual needs"
]

right_points = [
    "✓ One team serves the whole company",
    "✓ Best practices shared across groups",
    "✓ Tools that work across departments"
]

# Left box
lbox = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(2.3), Inches(3.3), Inches(2.2))
lbox.fill.solid()
lbox.fill.fore_color.rgb = RgbColor(240, 253, 250)
lbox.line.color.rgb = SAGE
lbox.line.width = Pt(1)

for i, pt in enumerate(left_points):
    lt = slide3.shapes.add_textbox(Inches(0.7), Inches(2.5) + i * Inches(0.65), Inches(2.9), Inches(0.6))
    tf = lt.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = pt
    p.font.size = Pt(13)
    p.font.color.rgb = SLATE

# Right box
rbox = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.533), Inches(2.3), Inches(3.3), Inches(2.2))
rbox.fill.solid()
rbox.fill.fore_color.rgb = RgbColor(240, 248, 255)
rbox.line.color.rgb = SKY
rbox.line.width = Pt(1)

for i, pt in enumerate(right_points):
    rt = slide3.shapes.add_textbox(Inches(9.733), Inches(2.5) + i * Inches(0.65), Inches(2.9), Inches(0.6))
    tf = rt.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = pt
    p.font.size = Pt(13)
    p.font.color.rgb = SLATE

# "What We Need From You" callout - addresses unspoken concerns
concern_box = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(6.4), Inches(12.333), Inches(0.9))
concern_box.fill.solid()
concern_box.fill.fore_color.rgb = RgbColor(255, 250, 240)  # Warm cream
concern_box.line.color.rgb = RgbColor(205, 133, 63)  # Warm accent
concern_box.line.width = Pt(1.5)

concern_title = slide3.shapes.add_textbox(Inches(0.8), Inches(6.5), Inches(3), Inches(0.4))
tf = concern_title.text_frame
p = tf.paragraphs[0]
p.text = "What we need from you:"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = RgbColor(180, 100, 50)

concern_text = slide3.shapes.add_textbox(Inches(3.8), Inches(6.5), Inches(8.8), Inches(0.7))
tf = concern_text.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "A conversation or two to help us understand your work. That's it. We handle the technical building — you stay focused on your job."
p.font.size = Pt(14)
p.font.color.rgb = SLATE

add_slide_footer(slide3)

# ============================================
# SLIDE 4: What We Offer - Data & Dashboards Focus
# ============================================
slide4 = prs.slides.add_slide(prs.slide_layouts[6])

bg4 = slide4.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
bg4.fill.solid()
bg4.fill.fore_color.rgb = WHITE
bg4.line.fill.background()

header4 = slide4.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.3))
header4.fill.solid()
header4.fill.fore_color.rgb = NAVY
header4.line.fill.background()

title4 = slide4.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.8))
tf = title4.text_frame
p = tf.paragraphs[0]
p.text = "What We Offer"
p.font.size = Pt(36)
p.font.bold = True
p.font.color.rgb = WHITE

# Five outcome-focused service areas
services = [
    {
        "color": SKY,
        "icon": "👁️",
        "title": "See Your Data Clearly",
        "subtitle": "From scattered spreadsheets to a single source of truth",
        "points": [
            "Dashboards that show project status at a glance",
            "Maps and visuals that update automatically",
            "One place to see what's happening across teams"
        ]
    },
    {
        "color": TEAL,
        "icon": "🔍",
        "title": "Understand What's Happening",
        "subtitle": "Turn raw data into answers",
        "points": [
            "Spot trends and patterns you'd otherwise miss",
            "Find root causes when things slow down",
            "Deep-dive analysis when you need to dig in"
        ]
    },
    {
        "color": SAGE,
        "icon": "🔮",
        "title": "Predict What's Coming",
        "subtitle": "See around corners before problems arrive",
        "points": [
            "Forecast timelines (like permit approvals)",
            "Flag risks early so you can act",
            "Model scenarios to plan ahead"
        ]
    },
    {
        "color": RgbColor(147, 112, 165),  # Purple
        "icon": "🤖",
        "title": "Work Smarter with AI",
        "subtitle": "Intelligent tools that multiply your capacity",
        "points": [
            "Generate reports and documents automatically",
            "Ask questions and get instant answers from your data",
            "Automate routine tasks so you can focus on judgment calls",
            "Get smart alerts when something needs your attention"
        ]
    },
    {
        "color": RgbColor(205, 133, 63),  # Warm orange
        "icon": "🧭",
        "title": "Navigate AI Confidently",
        "subtitle": "Make smart choices in a fast-moving landscape",
        "points": [
            "Guidance on using ChatGPT and similar tools safely",
            "Evaluate which AI products are actually worth it",
            "Stay ahead with university and industry connections"
        ]
    }
]

# Layout: 3 cards on top row, 2 centered on bottom row
card_width = Inches(4.0)
card_height = Inches(2.5)
gap = Inches(0.2)

# Top row: 3 cards
top_start_x = Inches(0.5)
top_y = Inches(1.5)

# Bottom row: 2 cards, centered
bottom_start_x = Inches(2.55)  # Centers two cards
bottom_y = Inches(4.2)

for i, svc in enumerate(services):
    if i < 3:
        # Top row
        x = top_start_x + i * (card_width + gap)
        y = top_y
    else:
        # Bottom row
        x = bottom_start_x + (i - 3) * (card_width + gap)
        y = bottom_y

    # Card background
    card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_width, card_height)
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = svc["color"]
    card.line.width = Pt(2)

    # Color accent bar
    bar = slide4.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(0.15), card_height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = svc["color"]
    bar.line.fill.background()

    # Icon and title
    title_txt = slide4.shapes.add_textbox(x + Inches(0.35), y + Inches(0.15), card_width - Inches(0.5), Inches(0.5))
    tf = title_txt.text_frame
    p = tf.paragraphs[0]
    p.text = f"{svc['icon']}  {svc['title']}"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = svc["color"]

    # Subtitle
    sub_txt = slide4.shapes.add_textbox(x + Inches(0.35), y + Inches(0.6), card_width - Inches(0.5), Inches(0.4))
    tf = sub_txt.text_frame
    p = tf.paragraphs[0]
    p.text = svc["subtitle"]
    p.font.size = Pt(13)
    p.font.italic = True
    p.font.color.rgb = MED_GRAY

    # Bullet points
    bullets = slide4.shapes.add_textbox(x + Inches(0.35), y + Inches(1.05), card_width - Inches(0.5), Inches(1.5))
    tf = bullets.text_frame
    tf.word_wrap = True
    for j, pt in enumerate(svc["points"]):
        if j == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = "• " + pt
        p.font.size = Pt(12)
        p.font.color.rgb = SLATE
        p.space_after = Pt(3)

add_slide_footer(slide4)

# ============================================
# SLIDE 5: Types of Solutions We Build
# ============================================
slide5 = prs.slides.add_slide(prs.slide_layouts[6])

bg5 = slide5.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
bg5.fill.solid()
bg5.fill.fore_color.rgb = WHITE
bg5.line.fill.background()

header5 = slide5.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.3))
header5.fill.solid()
header5.fill.fore_color.rgb = NAVY
header5.line.fill.background()

title5 = slide5.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.8))
tf = title5.text_frame
p = tf.paragraphs[0]
p.text = "Types of Solutions We Build"
p.font.size = Pt(36)
p.font.bold = True
p.font.color.rgb = WHITE

# Subtitle
sub5 = slide5.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.5), Inches(0.5))
tf = sub5.text_frame
p = tf.paragraphs[0]
p.text = "A spectrum of tools — from simple visibility to intelligent automation"
p.font.size = Pt(16)
p.font.color.rgb = MED_GRAY

# Solution types - horizontal flow showing progression
solution_types = [
    {
        "icon": "📊",
        "title": "Dashboards &\nVisualizations",
        "desc": "See your data clearly",
        "examples": "Status trackers, maps, charts that update automatically",
        "color": SKY
    },
    {
        "icon": "🔗",
        "title": "Connected\nData Systems",
        "desc": "Bring scattered info together",
        "examples": "Link spreadsheets, databases, and documents in one place",
        "color": TEAL
    },
    {
        "icon": "📈",
        "title": "Analysis &\nInsights",
        "desc": "Find patterns and trends",
        "examples": "What's working? What's slowing down? Where to focus?",
        "color": SAGE
    },
    {
        "icon": "🔮",
        "title": "Predictions &\nForecasts",
        "desc": "See what's likely coming",
        "examples": "Timeline estimates, risk flags, outcome probabilities",
        "color": RgbColor(147, 112, 165)
    },
    {
        "icon": "🤖",
        "title": "AI-Powered\nTools",
        "desc": "Smart automation",
        "examples": "Document generation, Q&A assistants, intelligent workflows",
        "color": RgbColor(205, 133, 63)
    }
]

card_width = Inches(2.35)
card_height = Inches(4.2)
start_x = Inches(0.5)
start_y = Inches(2.1)
gap = Inches(0.2)

for i, sol in enumerate(solution_types):
    x = start_x + i * (card_width + gap)

    # Card background
    card = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, start_y, card_width, card_height)
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = sol["color"]
    card.line.width = Pt(2)

    # Color bar at top
    bar = slide5.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, start_y, card_width, Inches(0.15))
    bar.fill.solid()
    bar.fill.fore_color.rgb = sol["color"]
    bar.line.fill.background()

    # Icon
    icon_bg = slide5.shapes.add_shape(MSO_SHAPE.OVAL, x + card_width/2 - Inches(0.4), start_y + Inches(0.35), Inches(0.8), Inches(0.8))
    icon_bg.fill.solid()
    icon_bg.fill.fore_color.rgb = sol["color"]
    icon_bg.line.fill.background()

    icon_txt = slide5.shapes.add_textbox(x + card_width/2 - Inches(0.35), start_y + Inches(0.45), Inches(0.7), Inches(0.6))
    tf = icon_txt.text_frame
    p = tf.paragraphs[0]
    p.text = sol["icon"]
    p.font.size = Pt(28)
    p.alignment = PP_ALIGN.CENTER

    # Title
    t = slide5.shapes.add_textbox(x + Inches(0.1), start_y + Inches(1.25), card_width - Inches(0.2), Inches(0.8))
    tf = t.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = sol["title"]
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = sol["color"]
    p.alignment = PP_ALIGN.CENTER

    # Short description
    d = slide5.shapes.add_textbox(x + Inches(0.1), start_y + Inches(2.05), card_width - Inches(0.2), Inches(0.5))
    tf = d.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = sol["desc"]
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = SLATE
    p.alignment = PP_ALIGN.CENTER

    # Examples
    e = slide5.shapes.add_textbox(x + Inches(0.1), start_y + Inches(2.55), card_width - Inches(0.2), Inches(1.4))
    tf = e.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = sol["examples"]
    p.font.size = Pt(11)
    p.font.color.rgb = MED_GRAY
    p.alignment = PP_ALIGN.CENTER

    # Arrow between cards (except last)
    if i < len(solution_types) - 1:
        arrow = slide5.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
            x + card_width + Inches(0.02), start_y + Inches(0.8),
            Inches(0.16), Inches(0.25))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = RgbColor(200, 210, 220)
        arrow.line.fill.background()

# Bottom note
note5 = slide5.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(12.3), Inches(0.5))
tf = note5.text_frame
p = tf.paragraphs[0]
p.text = "Let's look at some real examples →"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = TEAL
p.alignment = PP_ALIGN.CENTER

add_slide_footer(slide5)

# ============================================
# SLIDE 6: How to Engage With Us
# ============================================
slide6 = prs.slides.add_slide(prs.slide_layouts[6])

bg6 = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
bg6.fill.solid()
bg6.fill.fore_color.rgb = WHITE
bg6.line.fill.background()

header6 = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.3))
header6.fill.solid()
header6.fill.fore_color.rgb = NAVY
header6.line.fill.background()

title6 = slide6.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.8))
tf = title6.text_frame
p = tf.paragraphs[0]
p.text = "How to Work With Us"
p.font.size = Pt(36)
p.font.bold = True
p.font.color.rgb = WHITE

# Process steps
steps = [
    ("1", "Start a Conversation", "Tell us about a challenge you're\nfacing or an idea you have"),
    ("2", "We Learn Your World", "We ask questions to understand\nyour work and what success means"),
    ("3", "Design Together", "We propose approaches and you\nhelp us make them practical"),
    ("4", "Build & Refine", "We create, you test, we improve\nuntil it works for you"),
]

step_width = Inches(2.7)
step_start = Inches(0.9)
step_gap = Inches(0.5)
step_y = Inches(1.7)

for i, (num, title_text, desc) in enumerate(steps):
    x = step_start + i * (step_width + step_gap)

    # Number circle
    circle = slide6.shapes.add_shape(MSO_SHAPE.OVAL, x + step_width/2 - Inches(0.35), step_y, Inches(0.7), Inches(0.7))
    circle.fill.solid()
    circle.fill.fore_color.rgb = TEAL if i % 2 == 0 else SKY
    circle.line.fill.background()

    num_txt = slide6.shapes.add_textbox(x + step_width/2 - Inches(0.3), step_y + Inches(0.12), Inches(0.6), Inches(0.5))
    tf = num_txt.text_frame
    p = tf.paragraphs[0]
    p.text = num
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

    # Title
    t = slide6.shapes.add_textbox(x, step_y + Inches(0.9), step_width, Inches(0.5))
    tf = t.text_frame
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.alignment = PP_ALIGN.CENTER

    # Description
    d = slide6.shapes.add_textbox(x, step_y + Inches(1.4), step_width, Inches(1))
    tf = d.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = desc
    p.font.size = Pt(12)
    p.font.color.rgb = SLATE
    p.alignment = PP_ALIGN.CENTER

    # Arrow
    if i < len(steps) - 1:
        arrow = slide6.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
            x + step_width + Inches(0.1), step_y + Inches(0.2),
            Inches(0.3), Inches(0.3))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = RgbColor(200, 215, 230)
        arrow.line.fill.background()

# Bottom CTA box
cta_box = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(4.4), Inches(12.133), Inches(2.6))
cta_box.fill.solid()
cta_box.fill.fore_color.rgb = RgbColor(240, 248, 255)
cta_box.line.color.rgb = SKY
cta_box.line.width = Pt(2)

cta_title = slide6.shapes.add_textbox(Inches(0.9), Inches(4.6), Inches(11.5), Inches(0.6))
tf = cta_title.text_frame
p = tf.paragraphs[0]
p.text = "No question is too small. No idea is too ambitious."
p.font.size = Pt(22)
p.font.bold = True
p.font.color.rgb = NAVY
p.alignment = PP_ALIGN.CENTER

cta_examples = slide6.shapes.add_textbox(Inches(0.9), Inches(5.2), Inches(11.5), Inches(1.5))
tf = cta_examples.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = '"Can you help me make sense of this spreadsheet?"'
p.font.size = Pt(14)
p.font.color.rgb = SLATE
p.alignment = PP_ALIGN.CENTER
p.space_after = Pt(4)

p2 = tf.add_paragraph()
p2.text = '"We need to track 500 projects and predict which permits will be delayed."'
p2.font.size = Pt(14)
p2.font.color.rgb = SLATE
p2.alignment = PP_ALIGN.CENTER
p2.space_after = Pt(4)

p3 = tf.add_paragraph()
p3.text = '"Is there an AI tool that could help with this?"'
p3.font.size = Pt(14)
p3.font.color.rgb = SLATE
p3.alignment = PP_ALIGN.CENTER
p3.space_after = Pt(12)

p4 = tf.add_paragraph()
p4.text = "We're here to help. The best solutions often start with a simple conversation."
p4.font.size = Pt(16)
p4.font.bold = True
p4.font.color.rgb = TEAL
p4.alignment = PP_ALIGN.CENTER

add_slide_footer(slide6)

# Save the presentation
prs.save('/home/user/all/Enterprise_Analytics_AI_Team.pptx')
print("Presentation saved to: /home/user/all/Enterprise_Analytics_AI_Team.pptx")
