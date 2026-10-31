import os
from fastapi import FastAPI, Depends, HTTPException, Query, Response
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session
from typing import List, Optional

from backend.app.database import get_db, init_db, Question, Measurement, TechnicalFix, ContentPage, CitationTarget, AuditLog
from backend.app.schemas import (
    QuestionResponse, QuestionCreate, MeasurementResponse,
    TechnicalFixBase, ContentPageBase, CitationTargetBase,
    ContentBriefRequest, ContentBriefResponse
)
from backend.app.engine import calculate_visibility_metrics, generate_weekly_report, generate_content_brief, log_action

# Initialize database
init_db()

app = FastAPI(
    title="IncuBrix AEO, SEO & GEO Visibility Engine API",
    description="High-performance backend API powering the search, answer engine, and AI assistant tracking engine for IncuBrix.",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/favicon.ico", include_in_schema=False)
def favicon():
    return Response(content=b"", media_type="image/x-icon", status_code=204)

@app.get("/")
def root():
    return {
        "status": "online",
        "service": "IncuBrix Visibility Engine",
        "version": "1.0.0",
        "endpoints": [
            "/api/questions",
            "/api/measurements",
            "/api/metrics",
            "/api/fixes",
            "/api/pages",
            "/api/citations",
            "/api/brief/generate",
            "/api/reports/generate"
        ]
    }

@app.get("/api/health")
def health_check(db: Session = Depends(get_db)):
    q_count = db.query(Question).count()
    return {"status": "HEALTHY", "tracked_questions": q_count}

@app.get("/api/questions", response_model=List[QuestionResponse])
def get_questions(
    limit: int = 100,
    market: Optional[str] = None,
    category: Optional[str] = None,
    db: Session = Depends(get_db)
):
    query = db.query(Question)
    if market:
        query = query.filter(Question.market == market)
    if category:
        query = query.filter(Question.category == category)
    return query.limit(limit).all()

@app.get("/api/measurements")
def get_measurements(day: str = "Day 2", db: Session = Depends(get_db)):
    measurements = db.query(Measurement).filter(Measurement.day == day).all()
    return measurements

@app.get("/api/metrics")
def get_metrics(day: str = "Day 2", db: Session = Depends(get_db)):
    return calculate_visibility_metrics(db, day)

@app.get("/api/fixes")
def get_fixes(db: Session = Depends(get_db)):
    return db.query(TechnicalFix).all()

@app.get("/api/pages")
def get_pages(db: Session = Depends(get_db)):
    return db.query(ContentPage).all()

@app.get("/api/citations")
def get_citations(db: Session = Depends(get_db)):
    return db.query(CitationTarget).all()

@app.post("/api/brief/generate", response_model=ContentBriefResponse)
def create_brief(req: ContentBriefRequest, db: Session = Depends(get_db)):
    brief = generate_content_brief(req.model_dump())
    log_action(db, "GENERATE_BRIEF", "SUCCESS", f"Generated brief for query: {req.creator_question}")
    return brief

@app.post("/api/reports/generate")
def create_report(week: int = Query(1), db: Session = Depends(get_db)):
    rep_path = generate_weekly_report(week)
    if not rep_path:
        raise HTTPException(status_code=500, detail="Report generation failed")
    return {"status": "SUCCESS", "report_path": rep_path}

@app.get("/api/audit-logs")
def get_audit_logs(limit: int = 50, db: Session = Depends(get_db)):
    return db.query(AuditLog).order_by(AuditLog.id.desc()).limit(limit).all()
