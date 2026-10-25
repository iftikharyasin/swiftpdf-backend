from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter
import tempfile
import os
import shutil
import uuid

app = FastAPI(title="SwiftPDF High-Fidelity Converter API")

# Allow local HTML files and your future domain to communicate with this server
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

def cleanup_temp_dir(temp_dir: str):
    """Safely delete the temporary directory after the file has been sent."""
    try:
        shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception as e:
        print(f"Error cleaning up {temp_dir}: {e}")

@app.post("/api/convert")
async def convert_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    # Create a unique temporary directory for this specific request
    temp_dir = tempfile.mkdtemp()
    
    # Clean the filename and ensure safe paths
    safe_filename = file.filename or "document.pdf"
    in_pdf_path = os.path.join(temp_dir, safe_filename)
    base_name = os.path.splitext(safe_filename)[0]
    out_docx_path = os.path.join(temp_dir, f"{base_name}-converted.docx")

    try:
        # 1. Save the uploaded PDF to disk
        with open(in_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # 2. Convert using pdf2docx
        cv = Converter(in_pdf_path)
        
        # ✅ FIX: multi_processing=False prevents memory crashes on Render's free tier (512MB RAM)
        # ✅ FIX: page_layout_analysis=True helps handle multi-column academic papers better
        cv.convert(
            out_docx_path, 
            start=0, 
            end=None, 
            multi_processing=False, 
            page_layout_analysis=True
        )
        cv.close()

        # 3. Verify the file was created successfully
        if not os.path.exists(out_docx_path) or os.path.getsize(out_docx_path) == 0:
            raise HTTPException(status_code=500, detail="Conversion produced an empty file.")

        # 4. Add background task to delete the temp folder after the download finishes
        background_tasks.add_task(cleanup_temp_dir, temp_dir)

        # 5. Return the DOCX file
        return FileResponse(
            path=out_docx_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=os.path.basename(out_docx_path)
        )
        
    except Exception as e:
        # If anything fails, clean up immediately
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

@app.get("/")
def health_check():
    return {"status": "online", "service": "SwiftPDF Backend is Running"}
