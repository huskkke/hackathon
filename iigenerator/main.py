import os
import uuid
import json
import asyncio
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, UploadFile, Form, HTTPException, BackgroundTasks
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from doc_parser import parse_pdf, parse_docx
from llm_service import generate_slides_data
from pptx_builder import build_presentation

# ─── Setup ────────────────────────────────────────────────────────────────────

BASE_DIR = Path(__file__).parent
UPLOADS_DIR = BASE_DIR / "uploads"
OUTPUTS_DIR = BASE_DIR / "outputs"
UPLOADS_DIR.mkdir(exist_ok=True)
OUTPUTS_DIR.mkdir(exist_ok=True)

app = FastAPI(title="AI Presentation Generator", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory job status tracker
job_status: dict[str, dict] = {}


# ─── API Routes ───────────────────────────────────────────────────────────────

@app.get("/")
async def root():
    return FileResponse(BASE_DIR / "static" / "index.html")


@app.post("/api/generate")
async def generate_presentation(
    background_tasks: BackgroundTasks,
    prompt: str = Form(...),
    num_slides: int = Form(8),
    style: str = Form("corporate"),
    tone: str = Form("professional"),
    rostelecom_token: Optional[str] = Form(None),
    file: Optional[UploadFile] = File(None),
):
    job_id = str(uuid.uuid4())
    job_status[job_id] = {"status": "processing", "progress": 0, "message": "Initializing..."}

    # Read uploaded file
    doc_content = ""
    if file and file.filename:
        file_bytes = await file.read()
        fname = file.filename.lower()
        try:
            if fname.endswith(".pdf"):
                doc_content = parse_pdf(file_bytes)
            elif fname.endswith(".docx"):
                doc_content = parse_docx(file_bytes)
        except Exception as e:
            job_status[job_id] = {"status": "error", "message": f"File parse error: {e}"}
            return JSONResponse({"job_id": job_id})

    # Run generation in background
    background_tasks.add_task(
        run_generation,
        job_id=job_id,
        prompt=prompt,
        doc_content=doc_content,
        num_slides=num_slides,
        style=style,
        tone=tone,
        rostelecom_token=rostelecom_token,
    )

    return JSONResponse({"job_id": job_id})


async def run_generation(job_id: str, prompt: str, doc_content: str,
                          num_slides: int, style: str, tone: str,
                          rostelecom_token: str):
    try:
        job_status[job_id]["message"] = "Analyzing content with AI..."
        job_status[job_id]["progress"] = 20

        # Generate slide structure via LLM
        slides_data = await generate_slides_data(
            prompt=prompt,
            doc_content=doc_content,
            num_slides=num_slides,
            style=style,
            tone=tone,
            rostelecom_token=rostelecom_token,
        )

        job_status[job_id]["progress"] = 70
        job_status[job_id]["message"] = "Building presentation..."

        # Build PPTX
        pptx_bytes = build_presentation(slides_data, style)

        # Save to outputs
        output_path = OUTPUTS_DIR / f"{job_id}.pptx"
        with open(output_path, "wb") as f:
            f.write(pptx_bytes)

        # Save slides data for preview
        preview_path = OUTPUTS_DIR / f"{job_id}.json"
        with open(preview_path, "w", encoding="utf-8") as f:
            json.dump(slides_data, f, ensure_ascii=False, indent=2)

        job_status[job_id] = {
            "status": "done",
            "progress": 100,
            "message": "Done!",
            "slides": slides_data,
            "download_url": f"/api/download/{job_id}",
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        job_status[job_id] = {
            "status": "error",
            "progress": 0,
            "message": str(e),
        }


@app.get("/api/status/{job_id}")
async def get_status(job_id: str):
    status = job_status.get(job_id)
    if not status:
        raise HTTPException(404, "Job not found")
    return JSONResponse(status)


@app.get("/api/download/{job_id}")
async def download_pptx(job_id: str):
    output_path = OUTPUTS_DIR / f"{job_id}.pptx"
    if not output_path.exists():
        raise HTTPException(404, "File not found")
    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="presentation.pptx",
    )


# ─── Static Files ─────────────────────────────────────────────────────────────
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=7860, reload=False)