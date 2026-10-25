from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter
from pypdf import PdfReader, PdfWriter
from PIL import Image
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

        # Apply specific image downsampling rules based on user choice
        if level == "extreme":
            img_quality = 20
            scale_factor = 0.5  # Shrink image size by 50%
        elif level == "medium":
            img_quality = 40
            scale_factor = 0.75 # Shrink image size by 25%
        else:
            img_quality = 70
            scale_factor = 1.0  # Keep original image size

        writer = PdfWriter(clone_from=in_pdf_path)

        # Loop through every page to find and compress large images
        for page in writer.pages:
            for img in page.images:
                try:
                    pil_img = img.image
                    
                    # Convert to RGB to prevent transparency/alpha channel crashes
                    if pil_img.mode in ("RGBA", "P", "CMYK"):
                        pil_img = pil_img.convert("RGB")

                    # Downsample (Resize) the image to save massive amounts of space
                    if scale_factor < 1.0:
                        width, height = pil_img.size
                        new_size = (int(width * scale_factor), int(height * scale_factor))
                        pil_img = pil_img.resize(new_size, Image.Resampling.LANCZOS)
                    
                    # Replace the heavy image with our new, tiny JPEG version
                    img.replace(pil_img, quality=img_quality)
                except Exception as img_err:
                    print(f"Skipping image compression: {img_err}")
                    pass # Ignore vector masks or unsupported graphics

        with open(out_pdf_path, "wb") as f:
            writer.write(f)

        # Failsafe: If the PDF was already heavily compressed and our tool made it bigger, 
        # we return the original file so the user never gets a worse result.
        if os.path.getsize(out_pdf_path) >= os.path.getsize(in_pdf_path):
            shutil.copy(in_pdf_path, out_pdf_path)

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
