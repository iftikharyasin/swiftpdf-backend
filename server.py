from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pymupdf
from docx import Document
from docx.shared import Pt
import tempfile
import os
import shutil
import traceback
import re

app = FastAPI(title="SwiftPDF High-Fidelity Converter API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

def cleanup_temp_dir(temp_dir: str):
    try:
        shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception:
        pass

def clean_text(text):
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', text)
    text = re.sub(r'\s+(?=[a-zA-Z])', ' ', text) 
    return " ".join(text.split())

@app.post("/api/convert")
async def convert_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    temp_dir = tempfile.mkdtemp()
    safe_filename = file.filename or "document.pdf"
    in_pdf_path = os.path.join(temp_dir, safe_filename)
    base_name = os.path.splitext(safe_filename)[0]
    out_docx_path = os.path.join(temp_dir, f"{base_name}-converted.docx")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = pymupdf.open(in_pdf_path)
        docx = Document()
        
        style = docx.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(11)

        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            
            # Uses the lightweight C-extractor, not the heavy AI models
            text = page.get_text("text")
            
            if text.strip():
                cleaned_text = clean_text(text)
                paragraphs = cleaned_text.split('\n\n')
                for para in paragraphs:
                    if para.strip():
                        docx.add_paragraph(para.strip())

            if page_num < len(doc) - 1:
                docx.add_page_break()

        docx.save(out_docx_path)
        doc.close()

        if not os.path.exists(out_docx_path) or os.path.getsize(out_docx_path) == 0:
            raise HTTPException(status_code=500, detail="Conversion produced an empty file.")

        background_tasks.add_task(cleanup_temp_dir, temp_dir)

        return FileResponse(
            out_docx_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=os.path.basename(out_docx_path)
        )
        
    except Exception as e:
        print("CRITICAL ERROR:")
        traceback.print_exc()
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

@app.get("/")
def health_check():
    return {"status": "online", "service": "SwiftPDF Backend is Running"}
