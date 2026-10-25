from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pymupdf  # ✅ FIXED: Using modern 'pymupdf' instead of deprecated 'fitz'
from docx import Document
from docx.shared import Pt
import tempfile
import os
import shutil
import traceback

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

@app.post("/api/convert")
async def convert_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    temp_dir = tempfile.mkdtemp()
    safe_filename = file.filename or "document.pdf"
    in_pdf_path = os.path.join(temp_dir, safe_filename)
    base_name = os.path.splitext(safe_filename)[0]
    out_docx_path = os.path.join(temp_dir, f"{base_name}-converted.docx")

    try:
        # Save the uploaded PDF
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Open PDF with PyMuPDF
        doc = pymupdf.open(in_pdf_path)
        
        # Create a new Word Document
        docx = Document()
        
        # Setup default font
        style = docx.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(11)

        # Iterate through pages
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            
            # Extract text blocks
            blocks = page.get_text("blocks")
            
            # Sort blocks top-to-bottom, then left-to-right
            blocks.sort(key=lambda b: (round(b[1] / 10), b[0])) 
            
            for block in blocks:
                # ✅ FIXED: block[6] == 0 ensures we only process text, not images
                if block[6] == 0:
                    text = block[4].strip()
                    if text:
                        cleaned_text = " ".join(text.split())
                        docx.add_paragraph(cleaned_text)
            
            # Add a page break after each page (except the last)
            if page_num < len(doc) - 1:
                docx.add_page_break()

        # Save the Word document
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
        # This prints the exact error to your Render logs if it fails again
        print("CRITICAL ERROR:")
        traceback.print_exc()
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

@app.get("/")
def health_check():
    return {"status": "online", "service": "SwiftPDF Backend is Running"}
