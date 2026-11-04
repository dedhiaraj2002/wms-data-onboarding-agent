"""WMS onboarding HTTP endpoints."""
from pathlib import Path
from typing import Annotated

from fastapi import Depends, FastAPI, File, Form, HTTPException, UploadFile
from starlette.concurrency import run_in_threadpool

from app.database import DEFAULT_DATABASE, ValidationError
from app.validation import MAX_UPLOAD_BYTES, validate_workbook

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
    try:
        if not file.filename or not file.filename.lower().endswith(".xlsx"):
            raise ValidationError("INVALID_FILE_FORMAT", "Upload must have a .xlsx filename.")
        content = await file.read(MAX_UPLOAD_BYTES + 1)
        # Run parsing and SQLite reads outside the async event loop.
        return await run_in_threadpool(validate_workbook, content, customer_id, database)
    except ValidationError as exc:
        raise HTTPException(status_code=exc.status, detail={"code": exc.code, "message": str(exc)}) from exc
    finally:
        await file.close()
