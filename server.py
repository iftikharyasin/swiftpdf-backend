from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter
from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF engine for enterprise compression
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
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Using pdf2docx for visual layout preservation
        cv = Converter(in_pdf_path)
        
        cv.convert(
            out_docx_path, 
            start=0, 
            end=None, 
            multi_processing=False, 
            page_layout_analysis=True
        )
        cv.close()

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


@app.post("/api/compress")
async def compress_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...), level: str = Form("medium")):
    temp_dir = tempfile.mkdtemp()
    safe_filename = file.filename or "document.pdf"
    in_pdf_path = os.path.join(temp_dir, safe_filename)
    out_pdf_path = os.path.join(temp_dir, f"compressed_{safe_filename}")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Open PDF with PyMuPDF's C-based engine
        doc = fitz.open(in_pdf_path)
        
        # Dynamically apply compression algorithms based on user selection
        if level == "extreme":
            # Maximum garbage collection, deep stream compression, remove duplicate elements
            doc.save(out_pdf_path, garbage=4, deflate=True, clean=True)
        elif level == "medium":
            # Standard garbage collection and deflation
            doc.save(out_pdf_path, garbage=3, deflate=True, clean=True)
        else:
            # "Less" - Only remove obvious dead objects, keep maximum quality
            doc.save(out_pdf_path, garbage=1, deflate=True, clean=False)
            
        doc.close()

        background_tasks.add_task(cleanup_temp_dir, temp_dir)

        return FileResponse(
            out_pdf_path,
            media_type="application/pdf",
            filename=os.path.basename(out_pdf_path)
        )
    except Exception as e:
        print("CRITICAL ERROR:")
        traceback.print_exc()
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Compression failed: {str(e)}")


@app.post("/api/protect")
async def protect_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...), password: str = Form(...)):
    temp_dir = tempfile.mkdtemp()
    safe_filename = file.filename or "document.pdf"
    in_pdf_path = os.path.join(temp_dir, safe_filename)
    out_pdf_path = os.path.join(temp_dir, f"protected_{safe_filename}")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Encrypt the PDF with the provided password using 256-bit AES
        reader = PdfReader(in_pdf_path)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        writer.encrypt(user_password=password, owner_password=password, algorithm="AES-256")

        with open(out_pdf_path, "wb") as f:
            writer.write(f)

        background_tasks.add_task(cleanup_temp_dir, temp_dir)

        return FileResponse(
            out_pdf_path,
            media_type="application/pdf",
            filename=os.path.basename(out_pdf_path)
        )
    except Exception as e:
        print("CRITICAL ERROR:")
        traceback.print_exc()
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Protection failed: {str(e)}")

@app.get("/")
def health_check():
    return {"status": "online", "service": "SwiftPDF Backend is Running"}
