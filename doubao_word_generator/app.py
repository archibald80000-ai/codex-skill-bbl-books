from __future__ import annotations

import shutil
import sys
import uuid
import zipfile
from pathlib import Path
from typing import Annotated

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

SERVICE_ROOT = Path(__file__).resolve().parent
OUTPUTS_DIR = SERVICE_ROOT / "outputs"
SCRIPTS_DIR = SERVICE_ROOT / "scripts"
OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)

if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from generate_docx import SUPPORTED_DOC_TYPES, generate_from_paths, parse_formats  # type: ignore  # noqa: E402

app = FastAPI(
    title="BBL-books Word Generator",
    version="0.1.0",
    description="Generate downloadable DOCX/PDF booklets from uploaded materials.",
)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok", "service": "bbl-books-word-generator"}


def _safe_upload_name(name: str) -> str:
    cleaned = Path(name or "upload.bin").name.replace("\x00", "")
    return cleaned or "upload.bin"


def _extract_zip(zip_path: Path, target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as archive:
        for member in archive.infolist():
            member_path = target_dir / member.filename
            resolved = member_path.resolve()
            if target_dir.resolve() not in resolved.parents and resolved != target_dir.resolve():
                raise HTTPException(status_code=400, detail="zip contains unsafe paths")
        archive.extractall(target_dir)
    return target_dir


@app.post("/generate")
async def generate(
    doc_type: Annotated[str, Form()],
    output_formats: Annotated[str, Form()],
    title: Annotated[str | None, Form()] = None,
    files: Annotated[list[UploadFile] | None, File()] = None,
    zip_file: Annotated[UploadFile | None, File()] = None,
    source_path: Annotated[str | None, Form()] = None,
) -> dict:
    if doc_type not in SUPPORTED_DOC_TYPES:
        raise HTTPException(status_code=400, detail=f"doc_type must be one of: {', '.join(sorted(SUPPORTED_DOC_TYPES))}")
    try:
        parse_formats(output_formats)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    job_id = uuid.uuid4().hex[:12]
    job_dir = OUTPUTS_DIR / job_id
    input_dir = job_dir / "input"
    input_dir.mkdir(parents=True, exist_ok=True)

    inputs: list[str] = []
    if source_path:
        source = Path(source_path)
        if not source.exists():
            raise HTTPException(status_code=400, detail="source_path does not exist")
        inputs.append(str(source))

    if files:
        for upload in files:
            target = input_dir / _safe_upload_name(upload.filename)
            with target.open("wb") as handle:
                shutil.copyfileobj(upload.file, handle)
            inputs.append(str(target))

    if zip_file:
        zip_target = input_dir / _safe_upload_name(zip_file.filename or "materials.zip")
        with zip_target.open("wb") as handle:
            shutil.copyfileobj(zip_file.file, handle)
        extracted = _extract_zip(zip_target, input_dir / "unzipped")
        inputs.append(str(extracted))

    if not inputs:
        raise HTTPException(status_code=400, detail="provide files, zip_file, or source_path")

    try:
        result = generate_from_paths(
            inputs=inputs,
            doc_type=doc_type,
            output_formats=output_formats,
            title=title,
            output_dir=job_dir,
            job_id=job_id,
        )
    except Exception as exc:  # Keep API response clear while the internal report/log can hold details later.
        raise HTTPException(status_code=500, detail=f"generation failed: {exc}") from exc

    downloads = [
        {
            "type": Path(path).suffix.lstrip(".").lower(),
            "url": f"/download/{job_id}/{Path(path).name}",
        }
        for path in result["files"]
    ]
    return {"job_id": job_id, "status": "done", "downloads": downloads}


@app.get("/download/{job_id}/{filename}")
def download(job_id: str, filename: str) -> FileResponse:
    job_dir = (OUTPUTS_DIR / job_id).resolve()
    target = (job_dir / Path(filename).name).resolve()
    if job_dir not in target.parents and target != job_dir:
        raise HTTPException(status_code=400, detail="invalid download path")
    if not target.exists() or not target.is_file():
        raise HTTPException(status_code=404, detail="file not found")
    return FileResponse(target, filename=target.name)
