from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter
import tempfile
import os
import shutil

app = FastAPI(title="SwiftPDF High-Fidelity Converter API")

# Allow your local HTML files to communicate with this server
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/convert/{engine_name}")
async def convert_pdf(engine_name: str, file: UploadFile = File(...)):
    temp_dir = tempfile.mkdtemp()
    in_pdf_path = os.path.join(temp_dir, file.filename or "input.pdf")
    base_name = os.path.splitext(file.filename or "converted")[0]
    out_docx_path = os.path.join(temp_dir, f"{base_name}-{engine_name}.docx")

    try:
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        cv = Converter(in_pdf_path)
        cv.convert(out_docx_path, start=0, end=None, multi_processing=True, cpu_count=4)
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
