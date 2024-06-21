from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Create a presentation object
prs = Presentation()

# Set slide size to widescreen 16:9
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)

def set_font(placeholder, font_size):
    for paragraph in placeholder.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
            paragraph.alignment = PP_ALIGN.LEFT

# Slide 1: Cover Page
slide_1 = prs.slides.add_slide(prs.slide_layouts[0])
title_1 = slide_1.shapes.title
subtitle_1 = slide_1.placeholders[1]

title_1.text = "Artificial Intelligence in Healthcare"
subtitle_1.text = "An Overview of AI Applications in Healthcare"

set_font(title_1, 32)
set_font(subtitle_1, 16)

# Slide 2: Table of Contents
slide_2 = prs.slides.add_slide(prs.slide_layouts[1])
title_2 = slide_2.shapes.title
content_2 = slide_2.placeholders[1]

title_2.text = "Table of Contents"
content_2.text = (
    "1. Introduction to AI in Healthcare\n"
    "2. Applications of AI in Diagnostics\n"
    "3. AI in Treatment and Patient Care\n"
    "4. Challenges and Ethical Considerations\n"
    "5. Future of AI in Healthcare"
)

set_font(title_2, 32)
set_font(content_2, 16)

# Slide 3: Content One
slide_3 = prs.slides.add_slide(prs.slide_layouts[1])
title_3 = slide_3.shapes.title
content_3 = slide_3.placeholders[1]

title_3.text = "Introduction to AI in Healthcare"
content_3.text = (
    "Artificial Intelligence (AI) has the potential to revolutionize healthcare. "
    "It encompasses a range of technologies that enable machines to sense, comprehend, act, and learn."
)

set_font(title_3, 32)
set_font(content_3, 16)

# Slide 4: Content Two
slide_4 = prs.slides.add_slide(prs.slide_layouts[1])
title_4 = slide_4.shapes.title
content_4 = slide_4.placeholders[1]

title_4.text = "Applications of AI in Diagnostics"
content_4.text = (
    "AI algorithms can analyze medical images, detect patterns, and assist in diagnosing diseases. "
    "This can lead to earlier and more accurate diagnoses, improving patient outcomes."
)

set_font(title_4, 32)
set_font(content_4, 16)

# Slide 5: Content Three
slide_5 = prs.slides.add_slide(prs.slide_layouts[1])
title_5 = slide_5.shapes.title
content_5 = slide_5.placeholders[1]

title_5.text = "AI in Treatment and Patient Care"
content_5.text = (
    "AI can optimize treatment plans, personalize patient care, and provide predictive analytics. "
    "These capabilities help in managing chronic diseases and improving overall patient care."
)

set_font(title_5, 32)
set_font(content_5, 16)

# Save the presentation
prs.save('AI_in_Healthcare_Presentation.pptx')
