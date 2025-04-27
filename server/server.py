from fastapi import FastAPI, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import os
import json
from docx import Document
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from openai import OpenAI

client = OpenAI(api_key="")

# Initialize FastAPI
app = FastAPI()

# Enable CORS for frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Upload directory
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.post("/upload")
async def upload_file(file: UploadFile, parse_images: bool = Form(False), theme: str = Form("")):
    """
    Accepts a Word document and converts it into a PowerPoint presentation.
    """
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as f:
        f.write(file.file.read())

    theme_data = json.loads(theme)
    pptx_file = generate_presentation(file_path, theme_data)
    return {"pptx_url": f"http://localhost:8000/download/{os.path.basename(pptx_file)}"}


def generate_presentation(docx_path, theme):
    """
    Reads the Word document, summarizes content, and creates slides.
    """
    doc = Document(docx_path)
    full_text = "\n\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip()])
    slides_data = generate_multiple_slides(full_text)

    prs = Presentation()
    for slide_data in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        title_shape = slide.shapes.title
        title_shape.text = slide_data["title"]

        for i, layout in enumerate(prs.slide_layouts):
            print(f"Layout {i}: {layout.name}")

        for paragraph in title_shape.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Inches(0.56)  

        content_placeholder = None
        for shape in slide.placeholders:
            if shape.placeholder_format.type == 2:  
                content_placeholder = shape
                break

        if content_placeholder:
            text_frame = content_placeholder.text_frame
            text_frame.clear()
            for point in slide_data["bullets"]:
                p = text_frame.add_paragraph()
                p.level = 0
                run.text = f"• {point}"
                run = p.add_run(point)
                run.font.size = Pt(28)
        else:
            print("進來這裡了")
            left = Inches(1)
            top = Inches(2)
            width = Inches(8)
            height = Inches(4)
            content_box = slide.shapes.add_textbox(left, top, width, height)
            text_frame = content_box.text_frame
            text_frame.clear()
            text_frame.word_wrap = True
            for point in slide_data["bullets"]:
                p = text_frame.add_paragraph()
                p.level = 0
                run = p.add_run()        
                run.text = f"• {point}"
                run.font.size = Pt(28)

    apply_theme(prs, theme)

    pptx_filename = os.path.splitext(os.path.basename(docx_path))[0] + ".pptx"
    pptx_path = os.path.join(UPLOAD_DIR, pptx_filename)
    prs.save(pptx_path)

    return pptx_path


def chunk_and_summarize(full_text, chunk_size=1000):
    """
    Splits the full document into chunks and summarizes each using OpenAI.
    """
    paragraphs = full_text.split("\n\n")
    chunks = []
    current = ""

    for para in paragraphs:
        if len(current) + len(para) < chunk_size:
            current += para + "\n\n"
        else:
            chunks.append(current.strip())
            current = para + "\n\n"
    if current:
        chunks.append(current.strip())

    summarized_bullets = []

    for chunk in chunks:
        prompt = (
            "Please summarize the following section into 3 ~ 5 concise bullet points:\n"
            f"\"\"\"{chunk}\"\"\""
        )

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ],
            temperature=0.5,
            max_tokens=500
        )
        content = response.choices[0].message.content
        bullets = [line.strip()[2:] for line in content.splitlines() if line.strip().startswith("- ")]
        summarized_bullets.extend(bullets)

    return summarized_bullets


def generate_multiple_slides(full_text):
    """
    Generates slide structure based on summarized bullet points.
    """
    summarized_bullets = chunk_and_summarize(full_text)

    prompt = (
        "Based on the following summarized points, generate a PowerPoint slide structure:\n"
        "Each slide should have a title and 3–5 bullet points.\n\n"
        "Points:\n"
        "- " + "\n- ".join(summarized_bullets) + "\n\n"
        "Format:\n"
        "Slide 1:\n"
        "Title: ...\n"
        "Bullets:\n"
        "- ...\n"
        "- ...\n\n"
        "Slide 2:\n"
        "Title: ...\n"
        "Bullets:\n"
        "- ...\n"
    )
    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7,
        max_tokens=1500
    )
    output = response.choices[0].message.content
    return parse_multiple_slides(output)


def parse_multiple_slides(output_text):
    """
    Parses GPT output into structured slide dictionaries.
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
    Applies background color, text color, and background image if provided.
    """
    text_color = RGBColor(
        int(theme["text"][1:3], 16),
        int(theme["text"][3:5], 16),
        int(theme["text"][5:7], 16),
    )

    background_image = theme.get("backgroundImage", None)
    background_path = None
    print("background_image" + str(background_image))
    if background_image:
        background_path = os.path.join(os.path.dirname(__file__), "backgroundImages", background_image)
    print("background_path: " + str(background_path))
    for slide in prs.slides:
        if background_path and os.path.exists(background_path):
            pic = slide.shapes.add_picture(background_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
            slide.shapes._spTree.remove(pic._element)
            slide.shapes._spTree.insert(2, pic._element)

        else:
            # Fallback to solid color background
            background_color = RGBColor(
                int(theme["background"][1:3], 16),
                int(theme["background"][3:5], 16),
                int(theme["background"][5:7], 16),
            )
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = background_color

        # Set text color
        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                for para in shape.text_frame.paragraphs:
                    para.font.color.rgb = text_color

@app.get("/download/{filename}")
async def download_pptx(filename: str):
    """
    Serves the generated PPTX file.
    """
    pptx_path = os.path.join(UPLOAD_DIR, filename)
    if os.path.exists(pptx_path):
        return FileResponse(pptx_path, filename=filename)
    return {"error": "File not found"}