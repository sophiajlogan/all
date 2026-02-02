#!/usr/bin/env python3
"""
Create a PDF version of the Levantine Arabic presentation
"""

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# Try to register Arabic-supporting font
try:
    pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
    FONT_NAME = 'DejaVu'
except:
    FONT_NAME = 'Helvetica'

def create_pdf():
    doc = SimpleDocTemplate(
        "/home/user/all/Levantine_Arabic_Verb_Conjugation.pdf",
        pagesize=landscape(letter),
        rightMargin=0.5*inch,
        leftMargin=0.5*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch
    )

    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=32,
        textColor=colors.HexColor('#1a472a'),
        alignment=TA_CENTER,
        spaceAfter=20,
        fontName=FONT_NAME
    )

    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Normal'],
        fontSize=18,
        textColor=colors.HexColor('#4a4a4a'),
        alignment=TA_CENTER,
        spaceAfter=30,
        fontName=FONT_NAME
    )

    slide_title_style = ParagraphStyle(
        'SlideTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#1a472a'),
        spaceAfter=15,
        fontName=FONT_NAME
    )

    bullet_style = ParagraphStyle(
        'Bullet',
        parent=styles['Normal'],
        fontSize=14,
        textColor=colors.HexColor('#333333'),
        leftIndent=20,
        spaceAfter=8,
        fontName=FONT_NAME
    )

    note_style = ParagraphStyle(
        'Note',
        parent=styles['Normal'],
        fontSize=11,
        textColor=colors.HexColor('#666666'),
        fontName=FONT_NAME,
        spaceAfter=10
    )

    story = []

    # Slide 1: Title
    story.append(Spacer(1, 2*inch))
    story.append(Paragraph("Levantine Arabic Verb Conjugation", title_style))
    story.append(Paragraph("Past, Present &amp; Future Tenses with Modifiers", subtitle_style))
    story.append(Paragraph("العربية الشامية", subtitle_style))
    story.append(PageBreak())

    # Slide 2: What is Levantine Arabic
    story.append(Paragraph("What is Levantine Arabic?", slide_title_style))
    bullets = [
        "Spoken dialect in Syria, Lebanon, Jordan, and Palestine",
        "Different from Modern Standard Arabic (MSA) - the everyday spoken language",
        "Simplified verb system compared to MSA",
        "No dual form (only singular and plural)",
        "No grammatical case endings",
        "Rich in regional variations but mutually intelligible",
        "Uses a root-and-pattern system like all Arabic varieties"
    ]
    for b in bullets:
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(PageBreak())

    # Slide 3: Verb Basics
    story.append(Paragraph("Levantine Arabic Verb Basics", slide_title_style))
    bullets = [
        "Most verbs built on a 3-consonant root (e.g., k-t-b = write)",
        "Conjugation changes based on: person, number, gender, and tense",
        "Two main tenses: Past (perfect) and Present (imperfect)",
        "Future formed by adding prefix to present tense",
        "Subject pronouns often dropped (verb shows the subject)",
        "Example root: ك-ت-ب (k-t-b) meaning 'to write'"
    ]
    for b in bullets:
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(PageBreak())

    # Slide 4: Subject Pronouns
    story.append(Paragraph("Subject Pronouns Reference", slide_title_style))
    data = [
        ["Person", "Arabic", "Transliteration", "English"],
        ["1st Singular", "أنا", "ana", "I"],
        ["2nd Sing. Masc.", "إنت", "inta", "you (m)"],
        ["2nd Sing. Fem.", "إنتي", "inti", "you (f)"],
        ["3rd Sing. Masc.", "هو", "huwwe", "he"],
        ["3rd Sing. Fem.", "هي", "hiyye", "she"],
        ["1st Plural", "نحنا", "niḥna", "we"],
        ["2nd Plural", "إنتو", "intu", "you (pl)"],
        ["3rd Plural", "هم", "hunne", "they"],
    ]
    t = Table(data, colWidths=[2*inch, 1.5*inch, 1.5*inch, 1.5*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a472a')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('FONTSIZE', (0, 1), (-1, -1), 11),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cccccc')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f0f7f0'), colors.white]),
    ]))
    story.append(t)
    story.append(PageBreak())

    # Slide 5: Past Tense Introduction
    story.append(Paragraph("Past Tense (الماضي)", slide_title_style))
    bullets = [
        "Used for completed actions in the past",
        "Formed by adding suffixes to the verb stem",
        "The stem itself doesn't change - only the endings",
        "Example verb: كتب (katab) - 'to write'",
        "Past tense is the 'base' form of the verb",
        "Suffixes indicate person, number, and gender"
    ]
    for b in bullets:
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(PageBreak())

    # Slide 6: Past Tense Table
    story.append(Paragraph("Past Tense Conjugation", slide_title_style))
    data = [
        ["Person", "Suffix", "كتب (write)", "Transliteration", "Meaning"],
        ["I", "-t", "كتبت", "katabt", "I wrote"],
        ["you (m)", "-t", "كتبت", "katabt", "you wrote (m)"],
        ["you (f)", "-ti", "كتبتي", "katabti", "you wrote (f)"],
        ["he", "-Ø", "كتب", "katab", "he wrote"],
        ["she", "-et", "كتبت", "katabet", "she wrote"],
        ["we", "-na", "كتبنا", "katabna", "we wrote"],
        ["you (pl)", "-tu", "كتبتو", "katabtu", "you wrote (pl)"],
        ["they", "-u", "كتبو", "katabu", "they wrote"],
    ]
    t = Table(data, colWidths=[1.2*inch, 0.8*inch, 1.2*inch, 1.5*inch, 1.5*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a472a')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cccccc')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f0f7f0'), colors.white]),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph("Note: 'I' and 'you (m)' have the same form - context clarifies meaning", note_style))
    story.append(PageBreak())

    # Slide 7: Present Tense Introduction
    story.append(Paragraph("Present Tense (المضارع)", slide_title_style))
    bullets = [
        "Used for ongoing actions, habits, and general truths",
        "Formed with prefixes (and sometimes suffixes)",
        "The verb stem changes - vowel pattern differs from past",
        "Prefix 'b-' (bi-) added for indicative present in Levantine",
        "Without 'b-': subjunctive mood (after certain particles)",
        "Example: بكتب (baktub) - 'I write / I am writing'"
    ]
    for b in bullets:
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(PageBreak())

    # Slide 8: Present Tense Table
    story.append(Paragraph("Present Tense Conjugation", slide_title_style))
    data = [
        ["Person", "Prefix", "كتب (write)", "Transliteration", "Meaning"],
        ["I", "b-", "بكتب", "baktub", "I write"],
        ["you (m)", "bt-", "بتكتب", "btiktub", "you write (m)"],
        ["you (f)", "bt-...-i", "بتكتبي", "btiktubi", "you write (f)"],
        ["he", "bi-", "بيكتب", "byiktub", "he writes"],
        ["she", "bt-", "بتكتب", "btiktub", "she writes"],
        ["we", "mn-", "منكتب", "mniktub", "we write"],
        ["you (pl)", "bt-...-u", "بتكتبو", "btiktubu", "you write (pl)"],
        ["they", "bi-...-u", "بيكتبو", "byiktubu", "they write"],
    ]
    t = Table(data, colWidths=[1.2*inch, 0.9*inch, 1.2*inch, 1.4*inch, 1.5*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a472a')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cccccc')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f0f7f0'), colors.white]),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph("Note: 'you (m)' and 'she' have the same form", note_style))
    story.append(PageBreak())

    # Slide 9: Future Tense Introduction
    story.append(Paragraph("Future Tense (المستقبل)", slide_title_style))
    bullets = [
        "Formed by adding a prefix to the present tense (without 'b-')",
        "Main future marker: رح (raḥ) or حـ (ḥa-)",
        "رح is more common in Levantine dialects",
        "Can also use لح (laḥ) in some regions",
        "The verb conjugation follows present tense patterns",
        "Example: رح أكتب (raḥ aktub) - 'I will write'"
    ]
    for b in bullets:
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(PageBreak())

    # Slide 10: Future Tense Table
    story.append(Paragraph("Future Tense Conjugation", slide_title_style))
    data = [
        ["Person", "Future Form", "Transliteration", "Meaning"],
        ["I", "رح أكتب", "raḥ aktub", "I will write"],
        ["you (m)", "رح تكتب", "raḥ tiktub", "you will write (m)"],
        ["you (f)", "رح تكتبي", "raḥ tiktubi", "you will write (f)"],
        ["he", "رح يكتب", "raḥ yiktub", "he will write"],
        ["she", "رح تكتب", "raḥ tiktub", "she will write"],
        ["we", "رح نكتب", "raḥ niktub", "we will write"],
        ["you (pl)", "رح تكتبو", "raḥ tiktubu", "you will write (pl)"],
        ["they", "رح يكتبو", "raḥ yiktubu", "they will write"],
    ]
    t = Table(data, colWidths=[1.2*inch, 1.5*inch, 1.5*inch, 2*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a472a')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cccccc')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f0f7f0'), colors.white]),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph("Note: رح (raḥ) stays the same for all persons", note_style))
    story.append(PageBreak())

    # Slide 11: Modifiers Introduction
    story.append(Paragraph("Modifiers in Levantine Arabic", slide_title_style))
    bullets = [
        "Modifiers change the meaning or mood of verbs",
        "Key modifiers include:",
        "    - Negation (making verbs negative)",
        "    - Question formation",
        "    - Continuous/Progressive aspect",
        "    - Habitual aspect",
        "    - Imperative (commands)",
        "Each modifier has specific rules for each tense"
    ]
    for b in bullets:
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(PageBreak())

    # Slide 12: Negation Table
    story.append(Paragraph("Negation (النفي)", slide_title_style))
    data = [
        ["Tense", "Negation Pattern", "Example", "Meaning"],
        ["Past", "ما + verb", "ما كتبت", "I didn't write"],
        ["Present", "ما + verb (no b-)", "ما بكتب", "I don't write"],
        ["Present (alt)", "مش + b-verb", "مش بكتب", "I don't write"],
        ["Future", "مش رح + verb", "مش رح أكتب", "I won't write"],
        ["Future (alt)", "ما رح + verb", "ما رح أكتب", "I won't write"],
    ]
    t = Table(data, colWidths=[1.5*inch, 1.8*inch, 1.5*inch, 1.8*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a472a')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cccccc')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f0f7f0'), colors.white]),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph("ما (ma) and مش (mish) are both used for negation", note_style))
    story.append(PageBreak())

    # Slide 13: Negation Details
    story.append(Paragraph("Negation Details", slide_title_style))
    bullets = [
        "ما (ma) - General negation particle",
        "    Past: ما كتبت (ma katabt) - 'I didn't write'",
        "    Present: ما بكتب (ma baktub) - 'I don't write'",
        "مش (mish) - Also means 'not', often for emphasis",
        "    مش رح أكتب (mish raḥ aktub) - 'I will not write'",
        "In some dialects: ما...ش wraps around verb",
        "    ما كتبتش (ma katabtish) - 'I didn't write' (Egyptian influence)"
    ]
    for b in bullets:
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(PageBreak())

    # Slide 14: Question Formation
    story.append(Paragraph("Question Formation (الاستفهام)", slide_title_style))
    bullets = [
        "Yes/No questions: Rising intonation (no word change needed)",
        "    كتبت؟ (katabt?) - 'Did you write?'",
        "Question words placed at beginning:",
        "    شو (shu) - what: شو كتبت؟ 'What did you write?'",
        "    ليش (lesh) - why: ليش كتبت؟ 'Why did you write?'",
        "    كيف (kif) - how: كيف كتبت؟ 'How did you write?'",
        "    وين (wen) - where: وين كتبت؟ 'Where did you write?'",
        "    إيمتى (emta) - when: إيمتى كتبت؟ 'When did you write?'"
    ]
    for b in bullets:
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(PageBreak())

    # Slide 15: Progressive
    story.append(Paragraph("Progressive Aspect (عم)", slide_title_style))
    bullets = [
        "عم (ʿam) indicates ongoing action 'right now'",
        "Added before present tense verb (without b-)",
        "Similar to English '-ing' forms",
        "Examples:",
        "    عم بكتب (ʿam baktub) - 'I am writing (right now)'",
        "    عم يكتب (ʿam yiktub) - 'He is writing'",
        "Can combine with كان for past progressive:",
        "    كنت عم بكتب (kunt ʿam baktub) - 'I was writing'"
    ]
    for b in bullets:
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(PageBreak())

    # Slide 16: Using kaan
    story.append(Paragraph("Using كان (was/were)", slide_title_style))
    data = [
        ["Structure", "Example", "Transliteration", "Meaning"],
        ["كان + past", "كان كتب", "kān katab", "he had written"],
        ["كان + b-present", "كان بيكتب", "kān byiktub", "he used to write"],
        ["كان + عم + present", "كان عم يكتب", "kān ʿam yiktub", "he was writing"],
        ["كان + رح + present", "كان رح يكتب", "kān raḥ yiktub", "he was going to write"],
    ]
    t = Table(data, colWidths=[1.8*inch, 1.5*inch, 1.8*inch, 2*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a472a')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cccccc')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f0f7f0'), colors.white]),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph("كان conjugates: كنت (I was), كان (he was), كانت (she was), كنا (we were)...", note_style))
    story.append(PageBreak())

    # Slide 17: Imperative
    story.append(Paragraph("Imperative / Commands (الأمر)", slide_title_style))
    data = [
        ["Person", "Pattern", "Example (write!)", "Meaning"],
        ["you (m. sing)", "verb stem", "اكتب (ktub)", "Write!"],
        ["you (f. sing)", "verb stem + i", "اكتبي (ktubi)", "Write! (f)"],
        ["you (plural)", "verb stem + u", "اكتبو (ktubu)", "Write! (pl)"],
    ]
    t = Table(data, colWidths=[1.5*inch, 1.5*inch, 2*inch, 1.5*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a472a')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cccccc')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f0f7f0'), colors.white]),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph("Negative command: لا + present (no b-): لا تكتب (la tiktub) 'Don't write!'", note_style))
    story.append(PageBreak())

    # Slide 18: Common Verbs
    story.append(Paragraph("Common Verbs - All Tenses", slide_title_style))
    data = [
        ["Verb", "Past (he)", "Present (he)", "Future (he)"],
        ["to go", "راح (rāḥ)", "بيروح (birūḥ)", "رح يروح (raḥ yrūḥ)"],
        ["to eat", "أكل (akal)", "بياكل (byākul)", "رح ياكل (raḥ yākul)"],
        ["to see", "شاف (shāf)", "بيشوف (bishūf)", "رح يشوف (raḥ yshūf)"],
        ["to speak", "حكى (ḥaka)", "بيحكي (byiḥki)", "رح يحكي (raḥ yiḥki)"],
        ["to come", "إجا (ija)", "بيجي (bīji)", "رح يجي (raḥ yīji)"],
        ["to know", "عرف (ʿirif)", "بيعرف (byaʿrif)", "رح يعرف (raḥ yaʿrif)"],
        ["to want", "بدّو (biddo)*", "بدّو (biddo)", "—"],
    ]
    t = Table(data, colWidths=[1.2*inch, 1.8*inch, 1.8*inch, 2*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a472a')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cccccc')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f0f7f0'), colors.white]),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph("*بدّو is a pseudo-verb that conjugates differently: بدّي، بدّك، بدّو...", note_style))
    story.append(PageBreak())

    # Slide 19: Verb Patterns
    story.append(Paragraph("Verb Patterns (الأوزان)", slide_title_style))
    bullets = [
        "Levantine uses derived verb patterns (Forms I-X)",
        "Form I: Basic (فعل) - كتب (katab) 'write'",
        "Form II: Intensive/Causative (فعّل) - درّس (darras) 'teach'",
        "Form III: Reciprocal (فاعل) - ساعد (sāʿad) 'help'",
        "Form V: Reflexive of II (تفعّل) - تعلّم (tʿallam) 'learn'",
        "Form VI: Reciprocal (تفاعل) - تكلّم (tkallam) 'speak'",
        "Form VII: Passive (انفعل) - انكسر (nkasar) 'be broken'",
        "Each pattern conjugates similarly but with different vowels"
    ]
    for b in bullets:
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(PageBreak())

    # Slide 20: Regional Variations
    story.append(Paragraph("Regional Variations", slide_title_style))
    bullets = [
        "Syrian: Uses رح/لح for future; soft pronunciation",
        "Lebanese: Similar to Syrian; French loanwords common",
        "Palestinian: Uses رح; sometimes ح; glottal stops preserved",
        "Jordanian: Bedouin influences; قـ often becomes گ [g]",
        "All are mutually intelligible with minor differences",
        "This presentation uses general Levantine forms",
        "Local variations in vocabulary and some conjugations exist"
    ]
    for b in bullets:
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(PageBreak())

    # Slide 21: Quick Reference
    story.append(Paragraph("Quick Reference Summary", slide_title_style))
    data = [
        ["Tense", "Formation", "Example", "Negation"],
        ["Past", "Stem + suffix", "كتبت (katabt)", "ما كتبت"],
        ["Present", "b- + prefix + stem", "بكتب (baktub)", "ما بكتب"],
        ["Future", "رح + present (no b-)", "رح أكتب (raḥ aktub)", "مش رح أكتب"],
        ["Progressive", "عم + present", "عم بكتب (ʿam baktub)", "ما عم بكتب"],
        ["Command", "Stem only", "اكتب (ktub)", "لا تكتب"],
    ]
    t = Table(data, colWidths=[1.3*inch, 2*inch, 2*inch, 1.8*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a472a')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cccccc')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f0f7f0'), colors.white]),
    ]))
    story.append(t)
    story.append(PageBreak())

    # Slide 22: Tips
    story.append(Paragraph("Tips for Learning", slide_title_style))
    bullets = [
        "Master the pronouns and their corresponding verb endings first",
        "Learn verbs in groups with the same pattern",
        "Practice with common verbs: راح، أكل، شاف، حكى، جاب",
        "Listen to Levantine media (Syrian/Lebanese dramas, music)",
        "Focus on the 'b-' prefix pattern - it's essential for present tense",
        "Remember: Context often clarifies ambiguous forms",
        "Start with Form I verbs before learning derived forms",
        "بالتوفيق! (Good luck!)"
    ]
    for b in bullets:
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(PageBreak())

    # Slide 23: Thank you
    story.append(Spacer(1, 2*inch))
    story.append(Paragraph("!شكراً", title_style))
    story.append(Paragraph("Thank you for learning Levantine Arabic!", subtitle_style))
    story.append(Paragraph("Questions?", subtitle_style))

    doc.build(story)
    print("PDF saved as 'Levantine_Arabic_Verb_Conjugation.pdf'")

if __name__ == "__main__":
    create_pdf()
