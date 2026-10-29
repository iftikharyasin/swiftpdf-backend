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
import subprocess
import traceback
import pdfplumber
import pandas as pd

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


@app.post("/api/word-to-pdf")
async def word_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    temp_dir = tempfile.mkdtemp()
    safe_filename = file.filename or "document.docx"
    in_docx_path = os.path.join(temp_dir, safe_filename)
    base_name = os.path.splitext(safe_filename)[0]
    out_pdf_path = os.path.join(temp_dir, f"OptimaxPDF_WordToPDF_{base_name}.pdf")

    try:
        with open(in_docx_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Convert DOCX to PDF using LibreOffice with optimized layout filters
        process = subprocess.run(
            [
                "libreoffice", 
                "--headless", 
                "--convert-to", "pdf", 
                "--outdir", temp_dir, 
                in_docx_path
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        generated_pdf = os.path.join(temp_dir, f"{base_name}.pdf")
        if not os.path.exists(generated_pdf):
            for f in os.listdir(temp_dir):
                if f.endswith(".pdf"):
                    generated_pdf = os.path.join(temp_dir, f)
                    break

        if not os.path.exists(generated_pdf) or os.path.getsize(generated_pdf) == 0:
            raise HTTPException(status_code=500, detail="Word to PDF conversion produced an empty file.")

        shutil.move(generated_pdf, out_pdf_path)
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
        raise HTTPException(status_code=500, detail=f"Word to PDF conversion failed: {str(e)}")


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

        try:
            with pikepdf.open(in_pdf_path) as pdf:
                if pdf.is_encrypted:
                    raise ValueError("This file is already encrypted. Please unlock it first.")
                pdf.save(
                    out_pdf_path,
                    encryption=pikepdf.Encryption(
                        user=clean_password,
                        owner=clean_password,
                        R=6
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
async def unlock_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...), password: str = Form("")):
    temp_dir = tempfile.mkdtemp()
    safe_filename = file.filename or "document.pdf"
    in_pdf_path = os.path.join(temp_dir, safe_filename)
    out_pdf_path = os.path.join(temp_dir, f"OptimaxPDF_Unlocked_{safe_filename}")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        clean_pwd = password.strip() if password else ""
        unlocked_success = False

        passwords_to_try = []
        if clean_pwd:
            passwords_to_try.append(clean_pwd)
        if password and password != clean_pwd:
            passwords_to_try.append(password)
        if "" not in passwords_to_try:
            passwords_to_try.append("")

        # Tier 1: Try unlocking with pikepdf
        for pwd in passwords_to_try:
            try:
                with pikepdf.open(in_pdf_path, password=pwd) as pdf:
                    pdf.remove_security()
                    pdf.save(out_pdf_path)
                    unlocked_success = True
                    break
            except Exception:
                continue

        # Tier 2: Fallback to pypdf
        if not unlocked_success:
            try:
                reader = PdfReader(in_pdf_path)
                for pwd in passwords_to_try:
                    try:
                        if reader.is_encrypted:
                            res = reader.decrypt(pwd)
                        else:
                            res = 1
                        
                        if res > 0:
                            writer = PdfWriter()
                            for page in reader.pages:
                                writer.add_page(page)
                            with open(out_pdf_path, "wb") as f:
                                writer.write(f)
                            unlocked_success = True
                            break
                    except Exception:
                        continue
            except Exception as pypdf_err:
                print(f"pypdf fallback unlock error: {pypdf_err}")

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

        try:
            with pikepdf.open(in_pdf_path) as pdf_check:
                if pdf_check.is_encrypted:
                    raise ValueError("This file is password protected. Please unlock it first.")
        except pikepdf.PasswordError:
            raise ValueError("This file is password protected. Please unlock it first.")

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


# ---------------------------------------------------------
# NEW ROUTES: Excel to PDF, PPT to PDF, and PDF to Excel
# ---------------------------------------------------------

@app.post("/api/excel-to-pdf")
async def excel_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    temp_dir = tempfile.mkdtemp()
    safe_filename = file.filename or "document.xlsx"
    in_path = os.path.join(temp_dir, safe_filename)
    base_name = os.path.splitext(safe_filename)[0]
    out_pdf_path = os.path.join(temp_dir, f"OptimaxPDF_ExcelToPDF_{base_name}.pdf")

    try:
        with open(in_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Convert using LibreOffice
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", temp_dir, in_path],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE
        )
        
        # Robust file finding
        generated_pdf = os.path.join(temp_dir, f"{base_name}.pdf")
        if not os.path.exists(generated_pdf):
            for f in os.listdir(temp_dir):
                if f.endswith(".pdf"):
                    generated_pdf = os.path.join(temp_dir, f)
                    break
        
        if not os.path.exists(generated_pdf) or os.path.getsize(generated_pdf) == 0:
            raise HTTPException(status_code=500, detail="LibreOffice failed to generate PDF. Check file formatting.")

        shutil.move(generated_pdf, out_pdf_path)
        background_tasks.add_task(cleanup_temp_dir, temp_dir)
        
        return FileResponse(
            out_pdf_path, 
            media_type="application/pdf", 
            filename=os.path.basename(out_pdf_path)
        )
    except Exception as e:
        print("CRITICAL ERROR IN EXCEL TO PDF:")
        traceback.print_exc()
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")


@app.post("/api/ppt-to-pdf")
async def ppt_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    temp_dir = tempfile.mkdtemp()
    safe_filename = file.filename or "document.pptx"
    in_path = os.path.join(temp_dir, safe_filename)
    base_name = os.path.splitext(safe_filename)[0]
    out_pdf_path = os.path.join(temp_dir, f"OptimaxPDF_PPTtoPDF_{base_name}.pdf")

    try:
        with open(in_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Convert using LibreOffice
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", temp_dir, in_path],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE
        )
        
        # Robust file finding
        generated_pdf = os.path.join(temp_dir, f"{base_name}.pdf")
        if not os.path.exists(generated_pdf):
            for f in os.listdir(temp_dir):
                if f.endswith(".pdf"):
                    generated_pdf = os.path.join(temp_dir, f)
                    break
        
        if not os.path.exists(generated_pdf) or os.path.getsize(generated_pdf) == 0:
            raise HTTPException(status_code=500, detail="LibreOffice failed to generate PDF. Check file formatting.")

        shutil.move(generated_pdf, out_pdf_path)
        background_tasks.add_task(cleanup_temp_dir, temp_dir)
        
        return FileResponse(
            out_pdf_path, 
            media_type="application/pdf", 
            filename=os.path.basename(out_pdf_path)
        )
    except Exception as e:
        print("CRITICAL ERROR IN PPT TO PDF:")
        traceback.print_exc()
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")


@app.post("/api/pdf-to-excel")
async def pdf_to_excel(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    temp_dir = tempfile.mkdtemp()
    safe_filename = file.filename or "document.pdf"
    in_pdf_path = os.path.join(temp_dir, safe_filename)
    base_name = os.path.splitext(safe_filename)[0]
    out_xlsx_path = os.path.join(temp_dir, f"OptimaxPDF_Extracted_{base_name}.xlsx")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        all_tables = []
        with pdfplumber.open(in_pdf_path) as pdf:
            for page in pdf.pages:
                # 1. Try the default line-based extraction first
                tables = page.extract_tables()
                
                # 2. Fallback: If no borders exist, extract based on text spacing alignment
                if not tables:
                    tables = page.extract_tables({
                        "vertical_strategy": "text", 
                        "horizontal_strategy": "text",
                        "intersection_y_tolerance": 15
                    })

                for table in tables:
                    # Clean the table data (replace None with empty strings)
                    cleaned_table = [[cell if cell is not None else "" for cell in row] for row in table]
                    df = pd.DataFrame(cleaned_table)
                    all_tables.append(df)

        if not all_tables:
            raise ValueError("No tabular data detected in this PDF. Ensure the document contains structured columns.")

        # Write each table to a separate sheet in the Excel file
        with pd.ExcelWriter(out_xlsx_path, engine='openpyxl') as writer:
            for idx, df in enumerate(all_tables):
                df.to_excel(writer, sheet_name=f"Table_{idx+1}", index=False, header=False)

        background_tasks.add_task(cleanup_temp_dir, temp_dir)
        
        return FileResponse(
            out_xlsx_path, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            filename=os.path.basename(out_xlsx_path)
        )
    except ValueError as ve:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        print("CRITICAL ERROR IN PDF TO EXCEL:")
        traceback.print_exc()
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Table extraction failed: {str(e)}")


@app.get("/")
def health_check():
    return {"status": "online", "service": "Optimax PDF Backend is Running"}
