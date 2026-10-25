from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter
import tempfile
import os
import shutil

app = FastAPI(title="SwiftPDF High-Fidelity Converter API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ✅ Simplified endpoint (no engine_name needed)
@app.post("/api/convert")
async def convert_pdf(file: UploadFile = File(...)):
    temp_dir = tempfile.mkdtemp()
    in_pdf_path = os.path.join(temp_dir, file.filename or "input.pdf")
    base_name = os.path.splitext(file.filename or "converted")[0]
    out_docx_path = os.path.join(temp_dir, f"{base_name}-converted.docx")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        cv = Converter(in_pdf_path)
        # multi_processing=True can crash on Render's free tier (512MB RAM limit). 
        # Use False for stability.
        cv.convert(out_docx_path, start=0, end=None, multi_processing=False)
        cv.close()

        if not os.path.exists(out_docx_path) or os.path.getsize(out_docx_path) == 0:
            raise HTTPException(status_code=500, detail="Conversion produced an empty file.")

        return FileResponse(
            out_docx_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=os.path.basename(out_docx_path)
        )
    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/")
def health_check():
    return {"status": "online", "service": "SwiftPDF Backend is Running"}
