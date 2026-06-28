from __future__ import annotations

import os
from typing import Any, List, Optional

from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel, Field

from market_analysis_engine import build_report

APP_NAME = "Portfolio Link Market Backend"
BACKEND_TOKEN = os.getenv("BACKEND_API_TOKEN", "")

class AnalyzeRequest(BaseModel):
    headers: Optional[List[str]] = None
    rows: List[List[Any]] = Field(default_factory=list)

app = FastAPI(title=APP_NAME, version="0.1.0")

@app.get("/health")
def health():
    return {"ok": True, "service": APP_NAME}

@app.post("/analyze")
def analyze(req: AnalyzeRequest, x_backend_token: Optional[str] = Header(default=None)):
    if BACKEND_TOKEN and x_backend_token != BACKEND_TOKEN:
        raise HTTPException(status_code=401, detail="Invalid backend token")
    if not req.rows:
        raise HTTPException(status_code=400, detail="No Holdings rows received")
    try:
        return build_report(req.rows, headers=req.headers)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))
