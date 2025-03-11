from fastapi import FastAPI, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import os
import json
from docx import Document
from pptx import Presentation
from pptx.dml.color import RGBColor

# 初始化 FastAPI
app = FastAPI()

# 設定 CORS（允許前端請求）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],  # 允許 React 前端請求
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 設定上傳資料夾
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.post("/upload")
async def upload_file(file: UploadFile, parse_images: bool = Form(False), theme: str = Form("")):
    """ 接收上傳的 Word 文件並轉換為 PPTX """
    
    # 儲存 Word 檔案
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as f:
        f.write(file.file.read())

    # 解析 JSON 格式的主題設定
    theme_data = json.loads(theme)

    # 產生簡報
    pptx_file = generate_presentation(file_path, theme_data)

    # 返回下載連結
    return {"pptx_url": f"http://localhost:8000/download/{os.path.basename(pptx_file)}"}

def generate_presentation(docx_path, theme):
    """ 將 Word 內容轉換為 PowerPoint 簡報 """
    
    doc = Document(docx_path)
    prs = Presentation()

    for para in doc.paragraphs:
        if para.text.strip():
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            title = slide.shapes.title
            content = slide.placeholders[1]
            title.text = "Slide Title"
            content.text = para.text

    # 套用使用者選擇的主題
    apply_theme(prs, theme)

    # 儲存 PPTX
    pptx_filename = os.path.splitext(os.path.basename(docx_path))[0] + ".pptx"
    pptx_path = os.path.join(UPLOAD_DIR, pptx_filename)
    prs.save(pptx_path)
    
    return pptx_path

def apply_theme(prs, theme):
    """ 套用使用者選擇的簡報顏色主題 """

    # 解析 HEX 顏色為 RGB
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

    # 套用到所有投影片
    for slide in prs.slides:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = background_color

        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                for para in shape.text_frame.paragraphs:
                    para.font.color.rgb = text_color

@app.get("/download/{filename}")
async def download_pptx(filename: str):
    """ 提供下載已生成的 PowerPoint 簡報 """
    pptx_path = os.path.join(UPLOAD_DIR, filename)
    if os.path.exists(pptx_path):
        return FileResponse(pptx_path, filename=filename)
    return {"error": "File not found"}
