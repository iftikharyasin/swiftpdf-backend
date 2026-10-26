from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter
from pypdf import PdfReader, PdfWriter
from PIL import Image
import ocrmypdf
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

        if level == "extreme":
            img_quality = 20
            scale_factor = 0.5  
        elif level == "medium":
            img_quality = 40
            scale_factor = 0.75 
        else:
            img_quality = 70
            scale_factor = 1.0  

        writer = PdfWriter(clone_from=in_pdf_path)

        for page in writer.pages:
            for img in page.images:
                try:
                    pil_img = img.image
                    
                    if pil_img.mode in ("RGBA", "P", "CMYK"):
                        pil_img = pil_img.convert("RGB")

                    if scale_factor < 1.0:
                        width, height = pil_img.size
                        new_size = (int(width * scale_factor), int(height * scale_factor))
                        pil_img = pil_img.resize(new_size, Image.Resampling.LANCZOS)
                    
                    img.replace(pil_img, quality=img_quality)
                except Exception as img_err:
                    print(f"Skipping image compression: {img_err}")
                    pass 

        with open(out_pdf_path, "wb") as f:
            writer.write(f)

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
        
        if reader.is_encrypted:
            raise ValueError("This file is already encrypted. Please unlock it first.")
        
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
    except ValueError as ve:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        print("CRITICAL ERROR:")
        traceback.print_exc()
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Protection failed: {str(e)}")


@app.post("/api/unlock")
async def unlock_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...), password: str = Form(...)):
    temp_dir = tempfile.mkdtemp()
    safe_filename = file.filename or "document.pdf"
    in_pdf_path = os.path.join(temp_dir, safe_filename)
    out_pdf_path = os.path.join(temp_dir, f"unlocked_{safe_filename}")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        reader = PdfReader(in_pdf_path)
        
        if not reader.is_encrypted:
            raise ValueError("This PDF is not password protected.")

        try:
            decrypt_status = reader.decrypt(password)
            if decrypt_status == 0:
                raise ValueError("Incorrect password.")
        except Exception:
            raise ValueError("Incorrect password.")

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        with open(out_pdf_path, "wb") as f:
            writer.write(f)

        background_tasks.add_task(cleanup_temp_dir, temp_dir)

        return FileResponse(
            out_pdf_path,
            media_type="application/pdf",
            filename=os.path.basename(out_pdf_path)
        )
    except ValueError as ve:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=401, detail=str(ve))
    except Exception as e:
        print("CRITICAL ERROR:")
        traceback.print_exc()
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Unlock failed: {str(e)}")


@app.post("/api/ocr")
async def ocr_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...), language: str = Form("eng")):
    temp_dir = tempfile.mkdtemp()
    safe_filename = file.filename or "document.pdf"
    in_pdf_path = os.path.join(temp_dir, safe_filename)
    out_pdf_path = os.path.join(temp_dir, f"searchable_{safe_filename}")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Run OCRmyPDF engine with strict memory limits for Render Free Tier
        ocrmypdf.ocr(
            in_pdf_path, 
            out_pdf_path, 
            language=language, 
            force_ocr=True,
            optimize=0,              # Disable optimization to save RAM/CPU time
            jobs=1,                  # Force single job to prevent memory spikes
            max_ocr_image_mpixels=8, # Cap huge images at 8 megapixels
            skip_big=True            # Skip pages with massive images that would crash the server
        )

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
        raise HTTPException(status_code=500, detail=f"OCR Processing failed: {str(e)}")


@app.get("/")
def health_check():
    return {"status": "online", "service": "SwiftPDF Backend is Running"}
