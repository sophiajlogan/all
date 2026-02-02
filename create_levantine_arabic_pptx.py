#!/usr/bin/env python3
"""
Create a PowerPoint presentation about Levantine Arabic verb conjugation
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def add_title_slide(prs, title, subtitle):
    """Add a title slide"""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)

    # Add title shape
    left = Inches(0.5)
    top = Inches(2.5)
    width = Inches(9)
    height = Inches(1.5)
    title_box = slide.shapes.add_textbox(left, top, width, height)
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x1a, 0x47, 0x2a)  # Dark green
    p.alignment = PP_ALIGN.CENTER

    # Add subtitle
    left = Inches(0.5)
    top = Inches(4)
    width = Inches(9)
    height = Inches(1)
    subtitle_box = slide.shapes.add_textbox(left, top, width, height)
    tf = subtitle_box.text_frame
    p = tf.paragraphs[0]
    p.text = subtitle
    p.font.size = Pt(24)
    p.font.color.rgb = RGBColor(0x4a, 0x4a, 0x4a)
    p.alignment = PP_ALIGN.CENTER

    return slide

def add_content_slide(prs, title, content_items, rtl_examples=None):
    """Add a content slide with bullet points"""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)

    # Add title
    left = Inches(0.5)
    top = Inches(0.3)
    width = Inches(9)
    height = Inches(1)
    title_box = slide.shapes.add_textbox(left, top, width, height)
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x1a, 0x47, 0x2a)

    # Add content
    left = Inches(0.5)
    top = Inches(1.3)
    width = Inches(9)
    height = Inches(5.5)
    content_box = slide.shapes.add_textbox(left, top, width, height)
    tf = content_box.text_frame
    tf.word_wrap = True

    for i, item in enumerate(content_items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = "• " + item
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        p.space_after = Pt(12)

    return slide

def add_table_slide(prs, title, headers, rows, subtitle=None):
    """Add a slide with a table"""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)

    # Add title
    left = Inches(0.5)
    top = Inches(0.2)
    width = Inches(9)
    height = Inches(0.7)
    title_box = slide.shapes.add_textbox(left, top, width, height)
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x1a, 0x47, 0x2a)

    # Add subtitle if provided
    table_top = Inches(1.1)
    if subtitle:
        left = Inches(0.5)
        top = Inches(0.85)
        width = Inches(9)
        height = Inches(0.4)
        sub_box = slide.shapes.add_textbox(left, top, width, height)
        tf = sub_box.text_frame
        p = tf.paragraphs[0]
        p.text = subtitle
        p.font.size = Pt(16)
        p.font.italic = True
        p.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        table_top = Inches(1.3)

    # Create table
    num_cols = len(headers)
    num_rows = len(rows) + 1  # +1 for header

    left = Inches(0.5)
    width = Inches(9)
    height = Inches(0.4 * num_rows)

    table = slide.shapes.add_table(num_rows, num_cols, left, table_top, width, height).table

    # Set header row
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = header
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.size = Pt(14)
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xff, 0xff, 0xff)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x1a, 0x47, 0x2a)

    # Set data rows
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_data in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = cell_data
            cell.text_frame.paragraphs[0].font.size = Pt(13)
            if row_idx % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xf0, 0xf7, 0xf0)

    return slide

def create_presentation():
    """Create the full presentation"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Slide 1: Title
    add_title_slide(prs,
        "Levantine Arabic Verb Conjugation",
        "Past, Present & Future Tenses with Modifiers\nالعربية الشامية")

    # Slide 2: Introduction
    add_content_slide(prs, "What is Levantine Arabic?", [
        "Spoken dialect in Syria, Lebanon, Jordan, and Palestine",
        "Different from Modern Standard Arabic (MSA) - the everyday spoken language",
        "Simplified verb system compared to MSA",
        "No dual form (only singular and plural)",
        "No grammatical case endings",
        "Rich in regional variations but mutually intelligible",
        "Uses a root-and-pattern system like all Arabic varieties"
    ])

    # Slide 3: Verb Basics
    add_content_slide(prs, "Levantine Arabic Verb Basics", [
        "Most verbs built on a 3-consonant root (e.g., k-t-b = write)",
        "Conjugation changes based on: person, number, gender, and tense",
        "Two main tenses: Past (perfect) and Present (imperfect)",
        "Future formed by adding prefix to present tense",
        "Subject pronouns often dropped (verb shows the subject)",
        "Example root: ك-ت-ب (k-t-b) meaning 'to write'"
    ])

    # Slide 4: Subject Pronouns Reference
    add_table_slide(prs, "Subject Pronouns Reference",
        ["Person", "Arabic", "Transliteration", "English"],
        [
            ["1st Singular", "أنا", "ana", "I"],
            ["2nd Sing. Masc.", "إنت", "inta", "you (m)"],
            ["2nd Sing. Fem.", "إنتي", "inti", "you (f)"],
            ["3rd Sing. Masc.", "هو", "huwwe", "he"],
            ["3rd Sing. Fem.", "هي", "hiyye", "she"],
            ["1st Plural", "نحنا", "niḥna", "we"],
            ["2nd Plural", "إنتو", "intu", "you (pl)"],
            ["3rd Plural", "هم", "hunne", "they"],
        ])

    # Slide 5: Past Tense Introduction
    add_content_slide(prs, "Past Tense (الماضي)", [
        "Used for completed actions in the past",
        "Formed by adding suffixes to the verb stem",
        "The stem itself doesn't change - only the endings",
        "Example verb: كتب (katab) - 'to write'",
        "Past tense is the 'base' form of the verb",
        "Suffixes indicate person, number, and gender"
    ])

    # Slide 6: Past Tense Conjugation Table
    add_table_slide(prs, "Past Tense Conjugation",
        ["Person", "Suffix", "كتب (write)", "Transliteration", "Meaning"],
        [
            ["I", "-t", "كتبت", "katabt", "I wrote"],
            ["you (m)", "-t", "كتبت", "katabt", "you wrote (m)"],
            ["you (f)", "-ti", "كتبتي", "katabti", "you wrote (f)"],
            ["he", "-Ø", "كتب", "katab", "he wrote"],
            ["she", "-et", "كتبت", "katabet", "she wrote"],
            ["we", "-na", "كتبنا", "katabna", "we wrote"],
            ["you (pl)", "-tu", "كتبتو", "katabtu", "you wrote (pl)"],
            ["they", "-u", "كتبو", "katabu", "they wrote"],
        ],
        "Note: 'I' and 'you (m)' have the same form - context clarifies meaning")

    # Slide 7: Present Tense Introduction
    add_content_slide(prs, "Present Tense (المضارع)", [
        "Used for ongoing actions, habits, and general truths",
        "Formed with prefixes (and sometimes suffixes)",
        "The verb stem changes - vowel pattern differs from past",
        "Prefix 'b-' (bi-) added for indicative present in Levantine",
        "Without 'b-': subjunctive mood (after certain particles)",
        "Example: بكتب (baktub) - 'I write / I am writing'"
    ])

    # Slide 8: Present Tense Conjugation Table
    add_table_slide(prs, "Present Tense Conjugation",
        ["Person", "Prefix", "كتب (write)", "Transliteration", "Meaning"],
        [
            ["I", "b-", "بكتب", "baktub", "I write"],
            ["you (m)", "bt-", "بتكتب", "btiktub", "you write (m)"],
            ["you (f)", "bt-...-i", "بتكتبي", "btiktubi", "you write (f)"],
            ["he", "bi-", "بيكتب", "byiktub", "he writes"],
            ["she", "bt-", "بتكتب", "btiktub", "she writes"],
            ["we", "mn-", "منكتب", "mniktub", "we write"],
            ["you (pl)", "bt-...-u", "بتكتبو", "btiktubu", "you write (pl)"],
            ["they", "bi-...-u", "بيكتبو", "byiktubu", "they write"],
        ],
        "Note: 'you (m)' and 'she' have the same form")

    # Slide 9: Future Tense Introduction
    add_content_slide(prs, "Future Tense (المستقبل)", [
        "Formed by adding a prefix to the present tense (without 'b-')",
        "Main future marker: رح (raḥ) or حـ (ḥa-)",
        "رح is more common in Levantine dialects",
        "Can also use لح (laḥ) in some regions",
        "The verb conjugation follows present tense patterns",
        "Example: رح أكتب (raḥ aktub) - 'I will write'"
    ])

    # Slide 10: Future Tense Conjugation Table
    add_table_slide(prs, "Future Tense Conjugation",
        ["Person", "Future Form", "Transliteration", "Meaning"],
        [
            ["I", "رح أكتب", "raḥ aktub", "I will write"],
            ["you (m)", "رح تكتب", "raḥ tiktub", "you will write (m)"],
            ["you (f)", "رح تكتبي", "raḥ tiktubi", "you will write (f)"],
            ["he", "رح يكتب", "raḥ yiktub", "he will write"],
            ["she", "رح تكتب", "raḥ tiktub", "she will write"],
            ["we", "رح نكتب", "raḥ niktub", "we will write"],
            ["you (pl)", "رح تكتبو", "raḥ tiktubu", "you will write (pl)"],
            ["they", "رح يكتبو", "raḥ yiktubu", "they will write"],
        ],
        "Note: رح (raḥ) stays the same for all persons")

    # Slide 11: Modifiers Introduction
    add_content_slide(prs, "Modifiers in Levantine Arabic", [
        "Modifiers change the meaning or mood of verbs",
        "Key modifiers include:",
        "    - Negation (making verbs negative)",
        "    - Question formation",
        "    - Continuous/Progressive aspect",
        "    - Habitual aspect",
        "    - Imperative (commands)",
        "Each modifier has specific rules for each tense"
    ])

    # Slide 12: Negation
    add_table_slide(prs, "Negation (النفي)",
        ["Tense", "Negation Pattern", "Example", "Meaning"],
        [
            ["Past", "ما + verb", "ما كتبت", "I didn't write"],
            ["Present", "ما + verb (no b-)", "ما بكتب", "I don't write"],
            ["Present (alt)", "مش + b-verb", "مش بكتب", "I don't write"],
            ["Future", "مش رح + verb", "مش رح أكتب", "I won't write"],
            ["Future (alt)", "ما رح + verb", "ما رح أكتب", "I won't write"],
        ],
        "ما (ma) and مش (mish) are both used for negation")

    # Slide 13: Negation Details
    add_content_slide(prs, "Negation Details", [
        "ما (ma) - General negation particle",
        "    Past: ما كتبت (ma katabt) - 'I didn't write'",
        "    Present: ما بكتب (ma baktub) - 'I don't write'",
        "مش (mish) - Also means 'not', often for emphasis",
        "    مش رح أكتب (mish raḥ aktub) - 'I will not write'",
        "In some dialects: ما...ش wraps around verb",
        "    ما كتبتش (ma katabtish) - 'I didn't write' (Egyptian influence)"
    ])

    # Slide 14: Question Formation
    add_content_slide(prs, "Question Formation (الاستفهام)", [
        "Yes/No questions: Rising intonation (no word change needed)",
        "    كتبت؟ (katabt?) - 'Did you write?'",
        "Question words placed at beginning:",
        "    شو (shu) - what: شو كتبت؟ 'What did you write?'",
        "    ليش (lesh) - why: ليش كتبت؟ 'Why did you write?'",
        "    كيف (kif) - how: كيف كتبت؟ 'How did you write?'",
        "    وين (wen) - where: وين كتبت؟ 'Where did you write?'",
        "    إيمتى (emta) - when: إيمتى كتبت؟ 'When did you write?'"
    ])

    # Slide 15: Progressive/Continuous
    add_content_slide(prs, "Progressive Aspect (عم)", [
        "عم (ʿam) indicates ongoing action 'right now'",
        "Added before present tense verb (without b-)",
        "Similar to English '-ing' forms",
        "Examples:",
        "    عم بكتب (ʿam baktub) - 'I am writing (right now)'",
        "    عم يكتب (ʿam yiktub) - 'He is writing'",
        "Can combine with كان for past progressive:",
        "    كنت عم بكتب (kunt ʿam baktub) - 'I was writing'"
    ])

    # Slide 16: Verb with كان
    add_table_slide(prs, "Using كان (was/were)",
        ["Structure", "Example", "Transliteration", "Meaning"],
        [
            ["كان + past", "كان كتب", "kān katab", "he had written"],
            ["كان + b-present", "كان بيكتب", "kān byiktub", "he used to write"],
            ["كان + عم + present", "كان عم يكتب", "kān ʿam yiktub", "he was writing"],
            ["كان + رح + present", "كان رح يكتب", "kān raḥ yiktub", "he was going to write"],
        ],
        "كان conjugates: كنت (I was), كان (he was), كانت (she was), كنا (we were)...")

    # Slide 17: Imperative (Commands)
    add_table_slide(prs, "Imperative / Commands (الأمر)",
        ["Person", "Pattern", "Example (write!)", "Meaning"],
        [
            ["you (m. sing)", "verb stem", "اكتب (ktub)", "Write!"],
            ["you (f. sing)", "verb stem + i", "اكتبي (ktubi)", "Write! (f)"],
            ["you (plural)", "verb stem + u", "اكتبو (ktubu)", "Write! (pl)"],
        ],
        "Negative command: لا + present (no b-): لا تكتب (la tiktub) 'Don't write!'")

    # Slide 18: Common Verb Examples
    add_table_slide(prs, "Common Verbs - All Tenses",
        ["Verb", "Past (he)", "Present (he)", "Future (he)"],
        [
            ["to go", "راح (rāḥ)", "بيروح (birūḥ)", "رح يروح (raḥ yrūḥ)"],
            ["to eat", "أكل (akal)", "بياكل (byākul)", "رح ياكل (raḥ yākul)"],
            ["to see", "شاف (shāf)", "بيشوف (bishūf)", "رح يشوف (raḥ yshūf)"],
            ["to speak", "حكى (ḥaka)", "بيحكي (byiḥki)", "رح يحكي (raḥ yiḥki)"],
            ["to come", "إجا (ija)", "بيجي (bīji)", "رح يجي (raḥ yīji)"],
            ["to know", "عرف (ʿirif)", "بيعرف (byaʿrif)", "رح يعرف (raḥ yaʿrif)"],
            ["to want", "بدّو (biddo)*", "بدّو (biddo)", "—"],
        ],
        "*بدّو is a pseudo-verb that conjugates differently: بدّي، بدّك، بدّو...")

    # Slide 19: Verb Patterns
    add_content_slide(prs, "Verb Patterns (الأوزان)", [
        "Levantine uses derived verb patterns (Forms I-X)",
        "Form I: Basic (فعل) - كتب (katab) 'write'",
        "Form II: Intensive/Causative (فعّل) - درّس (darras) 'teach'",
        "Form III: Reciprocal (فاعل) - ساعد (sāʿad) 'help'",
        "Form V: Reflexive of II (تفعّل) - تعلّم (tʿallam) 'learn'",
        "Form VI: Reciprocal (تفاعل) - تكلّم (tkallam) 'speak'",
        "Form VII: Passive (انفعل) - انكسر (nkasar) 'be broken'",
        "Each pattern conjugates similarly but with different vowels"
    ])

    # Slide 20: Regional Variations
    add_content_slide(prs, "Regional Variations", [
        "Syrian: Uses رح/لح for future; soft pronunciation",
        "Lebanese: Similar to Syrian; French loanwords common",
        "Palestinian: Uses رح; sometimes ح; glottal stops preserved",
        "Jordanian: Bedouin influences; قـ often becomes گ [g]",
        "All are mutually intelligible with minor differences",
        "This presentation uses general Levantine forms",
        "Local variations in vocabulary and some conjugations exist"
    ])

    # Slide 21: Summary Table
    add_table_slide(prs, "Quick Reference Summary",
        ["Tense", "Formation", "Example", "Negation"],
        [
            ["Past", "Stem + suffix", "كتبت (katabt)", "ما كتبت"],
            ["Present", "b- + prefix + stem", "بكتب (baktub)", "ما بكتب"],
            ["Future", "رح + present (no b-)", "رح أكتب (raḥ aktub)", "مش رح أكتب"],
            ["Progressive", "عم + present", "عم بكتب (ʿam baktub)", "ما عم بكتب"],
            ["Command", "Stem only", "اكتب (ktub)", "لا تكتب"],
        ])

    # Slide 22: Practice Tips
    add_content_slide(prs, "Tips for Learning", [
        "Master the pronouns and their corresponding verb endings first",
        "Learn verbs in groups with the same pattern",
        "Practice with common verbs: راح، أكل، شاف، حكى، جاب",
        "Listen to Levantine media (Syrian/Lebanese dramas, music)",
        "Focus on the 'b-' prefix pattern - it's essential for present tense",
        "Remember: Context often clarifies ambiguous forms",
        "Start with Form I verbs before learning derived forms",
        "!بالتوفيق (Good luck!)"
    ])

    # Slide 23: Final slide
    add_title_slide(prs,
        "!شكراً",
        "Thank you for learning Levantine Arabic!\nQuestions?")

    # Save presentation
    prs.save('/home/user/all/Levantine_Arabic_Verb_Conjugation.pptx')
    print("Presentation saved as 'Levantine_Arabic_Verb_Conjugation.pptx'")

if __name__ == "__main__":
    create_presentation()
