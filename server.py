from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter
from pypdf import PdfReader, PdfWriter
from PIL import Image
import pikepdf
import ocrmypdf
import tempfile
import os
import shutil
import traceback

app = FastAPI(title="Optimax PDF High-Fidelity Cloud API")

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
    out_docx_path = os.path.join(temp_dir, f"OptimaxPDF_Converted_{base_name}.docx")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Check if the PDF is password-protected before running pdf2docx
        try:
            with pikepdf.open(in_pdf_path) as _:
                pass
        except pikepdf.PasswordError:
            raise ValueError("This file is password protected. Please unlock it first.")

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

    except ValueError as ve:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=403, detail=str(ve))
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
    out_pdf_path = os.path.join(temp_dir, f"OptimaxPDF_Compressed_{safe_filename}")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        reader = PdfReader(in_pdf_path)
        if reader.is_encrypted:
            raise ValueError("This file is password protected. Please unlock it first.")

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
    except ValueError as ve:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=403, detail=str(ve))
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
    out_pdf_path = os.path.join(temp_dir, f"OptimaxPDF_Protected_{safe_filename}")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        clean_password = password.strip()
        if not clean_password:
            raise ValueError("Password cannot be empty.")

        # Use pikepdf (QPDF) for rock-solid AES-256 encryption and accurate lock detection
        try:
            with pikepdf.open(in_pdf_path) as pdf:
                if pdf.is_encrypted:
                    raise ValueError("This file is already encrypted. Please unlock it first.")
                pdf.save(
                    out_pdf_path,
                    encryption=pikepdf.Encryption(
                        user=clean_password,
                        owner=clean_password,
                        R=6  # Enterprise AES-256 encryption
                    )
                )
        except pikepdf.PasswordError:
            raise ValueError("This file is already encrypted. Please unlock it first.")

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
    out_pdf_path = os.path.join(temp_dir, f"OptimaxPDF_Unlocked_{safe_filename}")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # 1. First check if the PDF can be opened without any password
        is_locked = False
        try:
            with pikepdf.open(in_pdf_path) as pdf_test:
                if not pdf_test.is_encrypted:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                    raise HTTPException(status_code=400, detail="This PDF is not password protected.")
                is_locked = True
        except pikepdf.PasswordError:
            is_locked = True

        # 2. Try unlocking with pikepdf (handles AES-256 R5/R6, AES-128, and RC4 natively)
        unlocked_success = False
        passwords_to_try = [password]
        if password.strip() != password:
            passwords_to_try.append(password.strip())

        for pwd in passwords_to_try:
            try:
                with pikepdf.open(in_pdf_path, password=pwd) as pdf:
                    pdf.save(out_pdf_path)
                    unlocked_success = True
                    break
            except pikepdf.PasswordError:
                continue

        # 3. Fallback to pypdf if pikepdf didn't unlock it
        if not unlocked_success:
            try:
                reader = PdfReader(in_pdf_path)
                for pwd in passwords_to_try:
                    if reader.decrypt(pwd) > 0:
                        writer = PdfWriter()
                        for page in reader.pages:
                            writer.add_page(page)
                        with open(out_pdf_path, "wb") as f:
                            writer.write(f)
                        unlocked_success = True
                        break
            except Exception as fallback_err:
                print(f"pypdf fallback error: {fallback_err}")

        if not unlocked_success:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise HTTPException(status_code=401, detail="Incorrect password. Please verify and try again.")

        background_tasks.add_task(cleanup_temp_dir, temp_dir)

        return FileResponse(
            out_pdf_path,
            media_type="application/pdf",
            filename=os.path.basename(out_pdf_path)
        )
    except HTTPException:
        raise
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
    out_pdf_path = os.path.join(temp_dir, f"OptimaxPDF_Searchable_{safe_filename}")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Check if the file is encrypted before attempting OCR
        try:
            with pikepdf.open(in_pdf_path) as pdf_check:
                if pdf_check.is_encrypted:
                    raise ValueError("This file is password protected. Please unlock it first.")
        except pikepdf.PasswordError:
            raise ValueError("This file is password protected. Please unlock it first.")

        # Run OCRmyPDF engine with strict memory limits for Render Free Tier
        ocrmypdf.ocr(
            in_pdf_path,
            out_pdf_path,
            language=language,
            force_ocr=True,
            optimize=0,
            jobs=1,
            skip_big=10.0
        )

        background_tasks.add_task(cleanup_temp_dir, temp_dir)

        return FileResponse(
            out_pdf_path,
            media_type="application/pdf",
            filename=os.path.basename(out_pdf_path)
        )
    except ValueError as ve:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=403, detail=str(ve))
    except Exception as e:
        print("CRITICAL ERROR:")
        traceback.print_exc()
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"OCR Processing failed: {str(e)}")


@app.get("/")
def health_check():
    return {"status": "online", "service": "Optimax PDF Backend is Running"}
