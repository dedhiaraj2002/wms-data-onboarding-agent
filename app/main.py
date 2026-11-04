"""Minimal WMS onboarding API."""
from fastapi import FastAPI

app = FastAPI(title="WMS Data Onboarding Agent")


@app.get("/health")
def health() -> dict[str, str]:
    """Report API liveness; this does not check external services."""
    return {"status": "ok"}
