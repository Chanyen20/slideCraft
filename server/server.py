from fastapi import FastAPI, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import os
import json
from docx import Document
from pptx import Presentation
from pptx.dml.color import RGBColor
from openai import OpenAI
import asyncio
from pptx.util import Inches

client = OpenAI(api_key="")

# Initialize FastAPI
app = FastAPI()

# Set up CORS to allow frontend requests
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"], 
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Set the upload directory
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.post("/upload")
async def upload_file(file: UploadFile, parse_images: bool = Form(False), theme: str = Form("")):
    """
    Accepts an uploaded Word file and converts it into a PowerPoint presentation.
    """
    # Save the uploaded Word file
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as f:
        f.write(file.file.read())

    # Parse theme data from the frontend (JSON format)
    theme_data = json.loads(theme)

    # Generate the presentation
    pptx_file = generate_presentation(file_path, theme_data)

    # Return the PPTX download URL
    return {"pptx_url": f"http://localhost:8000/download/{os.path.basename(pptx_file)}"}

def generate_presentation(docx_path, theme):
    """
    Converts the contents of a Word document into a PowerPoint presentation.
    """
    doc = Document(docx_path)
    full_text = "\n\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip()])
    slides_data = generate_multiple_slides(full_text)
    prs = Presentation()

    for slide_data in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title_shape = slide.shapes.title
        title_shape.text = slide_data["title"]

        left = Inches(1)
        top = Inches(2)
        width = Inches(8)
        height = Inches(4)
        content_box = slide.shapes.add_textbox(left, top, width, height)
        text_frame = content_box.text_frame
        text_frame.word_wrap = True

    for point in slide_data["bullets"]:
        p = text_frame.add_paragraph()
        p.text = point
        p.level = 0


    # Apply selected theme colors
    apply_theme(prs, theme)

    # Save the presentation
    pptx_filename = os.path.splitext(os.path.basename(docx_path))[0] + ".pptx"
    pptx_path = os.path.join(UPLOAD_DIR, pptx_filename)
    prs.save(pptx_path)
    
    return pptx_path

def generate_multiple_slides(full_text):
    """
    Uses OpenAI to generate multiple slide titles and bullet points from the full document text.
    """
    prompt = f"""
    You are a helpful assistant that generates a PowerPoint presentation from a Word document.

    Here is the full content:
    \"\"\"{full_text}\"\"\"

    Instructions:
    - Break it into logical sections for a slide deck.
    - For each slide:
    1. Title: 5-7 words starting with an action verb
    2. Bullet Points: 3-5 clear, concise points

    Format:
    Slide 1:
    Title: ...
    Bullets:
    - ...
    - ...

    Slide 2:
    Title: ...
    Bullets:
    - ...
    - ...
    """

    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You generate slide decks from long documents."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.7,
        max_tokens=1500
    )

    output = response.choices[0].message.content
    print("GPT Slide Output:\n", output)

    return parse_multiple_slides(output)

def parse_multiple_slides(output_text):
    """
    Parses the GPT output into a list of slide dicts.
    """
    slides = []
    current_slide = {}
    for line in output_text.splitlines():
        if line.lower().startswith("slide"):
            if current_slide:
                slides.append(current_slide)
            current_slide = {"title": "", "bullets": []}
        elif line.lower().startswith("title:"):
            current_slide["title"] = line.split(":", 1)[1].strip()
        elif line.strip().startswith("- "):
            current_slide["bullets"].append(line.strip()[2:])
    if current_slide:
        slides.append(current_slide)
    return slides

def apply_theme(prs, theme):
    """
    Applies the user's selected color theme to all slides.
    """
    background_color = RGBColor(
        int(theme["background"][1:3], 16),
        int(theme["background"][3:5], 16),
        int(theme["background"][5:7], 16),
    )
    text_color = RGBColor(
        int(theme["text"][1:3], 16),
        int(theme["text"][3:5], 16),
        int(theme["text"][5:7], 16),
    )

    for slide in prs.slides:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = background_color

        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                for para in shape.text_frame.paragraphs:
                    para.font.color.rgb = text_color

@app.get("/download/{filename}")
async def download_pptx(filename: str):
    """
    Provides a download link for the generated PowerPoint presentation.
    """
    pptx_path = os.path.join(UPLOAD_DIR, filename)
    if os.path.exists(pptx_path):
        return FileResponse(pptx_path, filename=filename)
    return {"error": "File not found"}