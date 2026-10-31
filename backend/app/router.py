from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session
from typing import List, Optional

from backend.app.database import get_db, Question, Measurement, TechnicalFix, ContentPage, CitationTarget, AuditLog
from backend.app.schemas import (
    QuestionResponse, MeasurementResponse,
    ContentBriefRequest, ContentBriefResponse
)
from backend.app.engine import calculate_visibility_metrics, generate_weekly_report, generate_content_brief, log_action

router = APIRouter(prefix="/api", tags=["Visibility Engine"])

@router.get("/health")
def health_check(db: Session = Depends(get_db)):
    q_count = db.query(Question).count()
    return {"status": "HEALTHY", "tracked_questions": q_count}

@router.get("/questions", response_model=List[QuestionResponse])
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

@router.get("/measurements")
def get_measurements(day: str = "Day 2", db: Session = Depends(get_db)):
    return db.query(Measurement).filter(Measurement.day == day).all()

@router.get("/metrics")
def get_metrics(day: str = "Day 2", db: Session = Depends(get_db)):
    return calculate_visibility_metrics(db, day)

@router.get("/fixes")
def get_fixes(db: Session = Depends(get_db)):
    return db.query(TechnicalFix).all()

@router.get("/pages")
def get_pages(db: Session = Depends(get_db)):
    return db.query(ContentPage).all()

@router.get("/citations")
def get_citations(db: Session = Depends(get_db)):
    return db.query(CitationTarget).all()

@router.post("/brief/generate", response_model=ContentBriefResponse)
def create_brief(req: ContentBriefRequest, db: Session = Depends(get_db)):
    brief = generate_content_brief(req.model_dump())
    log_action(db, "GENERATE_BRIEF", "SUCCESS", f"Generated brief for query: {req.creator_question}")
    return brief

@router.post("/reports/generate")
def create_report(week: int = Query(1), db: Session = Depends(get_db)):
    rep_path = generate_weekly_report(week)
    if not rep_path:
        raise HTTPException(status_code=500, detail="Report generation failed")
    return {"status": "SUCCESS", "report_path": rep_path}

@router.get("/audit-logs")
def get_audit_logs(limit: int = 50, db: Session = Depends(get_db)):
    return db.query(AuditLog).order_by(AuditLog.id.desc()).limit(limit).all()
