"""WMS onboarding HTTP endpoints."""
from pathlib import Path
from typing import Annotated, Literal
from functools import partial

from fastapi import Depends, FastAPI, File, Form, HTTPException, UploadFile
from starlette.concurrency import run_in_threadpool

from app.database import DEFAULT_DATABASE, ValidationError
from app.validation import MAX_UPLOAD_BYTES, validate_workbook
from app.investigation import investigate_workbook
from app.agent import analyze_workbook

app = FastAPI(title="WMS Data Onboarding Agent")


@app.get("/health")
def health() -> dict[str, str]:
    """Report API liveness; this does not check external services."""
    return {"status": "ok"}


def database_path() -> Path:
    return DEFAULT_DATABASE


@app.post("/validate")
async def validate(
    file: Annotated[UploadFile, File()],
    customer_id: Annotated[str, Form(min_length=1)],
    database: Annotated[Path, Depends(database_path)],
) -> dict:
    return await process_upload(file, customer_id, database, validate_workbook)


@app.post("/investigate")
async def investigate(
    file: Annotated[UploadFile, File()],
    customer_id: Annotated[str, Form(min_length=1)],
    database: Annotated[Path, Depends(database_path)],
) -> dict:
    return await process_upload(file, customer_id, database, investigate_workbook)


@app.post("/agent/analyze")
async def agent_analyze(
    file: Annotated[UploadFile, File()],
    customer_id: Annotated[str, Form(min_length=1)],
    database: Annotated[Path, Depends(database_path)],
    mode: Annotated[Literal["mock", "openai"], Form()] = "mock",
) -> dict:
    return await process_upload(file, customer_id, database, partial(analyze_workbook, mode=mode))


async def process_upload(file: UploadFile, customer_id: str, database: Path, operation) -> dict:
    """Share upload protections and error handling across both endpoints."""
    try:
        if not file.filename or not file.filename.lower().endswith(".xlsx"):
            raise ValidationError("INVALID_FILE_FORMAT", "Upload must have a .xlsx filename.")
        content = await file.read(MAX_UPLOAD_BYTES + 1)
        # Run parsing and SQLite reads outside the async event loop.
        return await run_in_threadpool(operation, content, customer_id, database)
    except ValidationError as exc:
        raise HTTPException(status_code=exc.status, detail={"code": exc.code, "message": str(exc)}) from exc
    finally:
        await file.close()
