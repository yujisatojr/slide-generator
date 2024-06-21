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

def set_title_position_and_font(slide, title_text):
    title = slide.shapes.title
    title.text = title_text
    title.left = Inches(0.5)
    title.top = Inches(0.5)
    title.width = Inches(12)
    title.height = Inches(0.6)
    set_font(title, 32)

# Slide 1: Cover Page
slide_1 = prs.slides.add_slide(prs.slide_layouts[6])

# Set background image with padding
padding = Inches(0.15)
background_image_path = './assets/pexels-pixabay-273250.jpg'
slide_1.shapes.add_picture(
    background_image_path, 
    padding, 
    padding, 
    prs.slide_width - 2 * padding, 
    prs.slide_height - 2 * padding
)

# Add title with specific position and size
title_1 = slide_1.shapes.add_textbox(Inches(0.49), Inches(2.72), prs.slide_width - Inches(1), Inches(1.5))
title_frame = title_1.text_frame
title_paragraph = title_frame.add_paragraph()
title_paragraph.text = "Artificial Intelligence in Healthcare"
title_paragraph.font.size = Pt(44)
title_paragraph.font.color.rgb = RGBColor(255, 255, 255)  # Set text color to white
title_paragraph.alignment = PP_ALIGN.LEFT

# Add subtitle with specific position and size
subtitle_1 = slide_1.shapes.add_textbox(Inches(0.49), Inches(3.76), prs.slide_width - Inches(1), Inches(1))
subtitle_frame = subtitle_1.text_frame
subtitle_paragraph = subtitle_frame.add_paragraph()
subtitle_paragraph.text = "An Overview of AI Applications in Healthcare"
subtitle_paragraph.font.size = Pt(20)
subtitle_paragraph.font.color.rgb = RGBColor(255, 255, 255)  # Set text color to white
subtitle_paragraph.alignment = PP_ALIGN.LEFT

# Slide 2: Table of Contents
slide_2 = prs.slides.add_slide(prs.slide_layouts[5])  # Using a blank slide layout to avoid placeholder

set_title_position_and_font(slide_2, "Table of Contents")

# Add a textbox for the content with specified position and size
content_2 = slide_2.shapes.add_textbox(Inches(0.5), Inches(1.5), prs.slide_width - Inches(1), prs.slide_height - Inches(2))
text_frame_2 = content_2.text_frame

# Add paragraphs for each item in the table of contents
toc_items = [
    "1. Introduction to AI in Healthcare",
    "2. Applications of AI in Diagnostics",
    "3. AI in Treatment and Patient Care",
    "4. Challenges and Ethical Considerations",
    "5. Future of AI in Healthcare"
]

for item in toc_items:
    p = text_frame_2.add_paragraph()
    p.text = item
    p.font.size = Pt(24)
    p.alignment = PP_ALIGN.LEFT

# Slide 3: Content One
slide_3 = prs.slides.add_slide(prs.slide_layouts[1])
set_title_position_and_font(slide_3, "Introduction to AI in Healthcare")
content_3 = slide_3.placeholders[1]

content_3.text = (
    "Artificial Intelligence (AI) has the potential to revolutionize healthcare. "
    "It encompasses a range of technologies that enable machines to sense, comprehend, act, and learn."
)
set_font(content_3, 16)

# Slide 4: Content Two
slide_4 = prs.slides.add_slide(prs.slide_layouts[1])
set_title_position_and_font(slide_4, "Applications of AI in Diagnostics")
content_4 = slide_4.placeholders[1]

content_4.text = (
    "AI algorithms can analyze medical images, detect patterns, and assist in diagnosing diseases. "
    "This can lead to earlier and more accurate diagnoses, improving patient outcomes."
)
set_font(content_4, 16)

# Slide 5: Content Three
slide_5 = prs.slides.add_slide(prs.slide_layouts[1])
set_title_position_and_font(slide_5, "AI in Treatment and Patient Care")
content_5 = slide_5.placeholders[1]

content_5.text = (
    "AI can optimize treatment plans, personalize patient care, and provide predictive analytics. "
    "These capabilities help in managing chronic diseases and improving overall patient care."
)
set_font(content_5, 16)

# Save the presentation
prs.save('AI_in_Healthcare_Presentation.pptx')
