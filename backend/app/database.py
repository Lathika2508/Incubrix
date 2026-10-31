import os
from sqlalchemy import create_engine, Column, Integer, String, Text, Boolean, Float, DateTime, ForeignKey
from sqlalchemy.orm import declarative_base, sessionmaker, relationship
import datetime

DB_PATH = r"D:\incubrix project final\visibility_engine.db"
DATABASE_URL = f"sqlite:///{DB_PATH}"

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class Question(Base):
    __tablename__ = "questions"
    id = Column(Integer, primary_key=True, index=True)
    question_id = Column(String(50), unique=True, index=True)
    question = Column(Text, nullable=False)
    category = Column(String(100), index=True)
    intent = Column(String(50), index=True)
    source_type = Column(String(100))
    source_url = Column(Text)
    market = Column(String(10), index=True)
    priority = Column(String(20))
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    measurements = relationship("Measurement", back_populates="question", cascade="all, delete-orphan")

class Measurement(Base):
    __tablename__ = "measurements"
    id = Column(Integer, primary_key=True, index=True)
    question_id = Column(String(50), ForeignKey("questions.question_id"), index=True)
    day = Column(String(20), index=True) # "Day 1", "Day 2"
    measurement_date = Column(String(30))
    google_position = Column(String(50))
    answer_box_present = Column(Boolean, default=False)
    quoted_source_url = Column(Text)
    chatgpt_mentioned = Column(Boolean, default=False)
    gemini_mentioned = Column(Boolean, default=False)
    perplexity_mentioned = Column(Boolean, default=False)
    claude_copilot_mentioned = Column(Boolean, default=False)
    competitors_named = Column(Text)
    raw_response_path = Column(Text)
    changed_from_previous = Column(Boolean, default=False)

    question = relationship("Question", back_populates="measurements")

class TechnicalFix(Base):
    __tablename__ = "technical_fixes"
    id = Column(Integer, primary_key=True, index=True)
    fix_id = Column(String(50), unique=True, index=True)
    page_url = Column(Text, nullable=False)
    problem = Column(Text, nullable=False)
    evidence_url = Column(Text)
    fix_summary = Column(Text)
    code_path = Column(Text)
    validator_used = Column(String(100))
    validated = Column(String(10), default="Yes")
    impact = Column(String(20))
    effort = Column(String(20))

class ContentPage(Base):
    __tablename__ = "content_pages"
    id = Column(Integer, primary_key=True, index=True)
    item_id = Column(String(50), unique=True, index=True)
    slug = Column(String(255))
    title = Column(Text)
    primary_query = Column(String(255))
    target_qids = Column(Text)
    draft_path = Column(Text)
    status = Column(String(50), default="Ready")

class CitationTarget(Base):
    __tablename__ = "citation_targets"
    id = Column(Integer, primary_key=True, index=True)
    target_id = Column(String(50), unique=True, index=True)
    website = Column(String(100))
    target_url = Column(Text)
    page_type = Column(String(100))
    classification = Column(String(50))
    priority = Column(String(20))
    mapped_qids = Column(Text)

class WeeklyReportModel(Base):
    __tablename__ = "weekly_reports"
    id = Column(Integer, primary_key=True, index=True)
    week_number = Column(Integer, index=True)
    report_date = Column(String(30))
    summary_markdown = Column(Text)
    google_presence_rate = Column(Float)
    ai_mention_rate = Column(Float)
    total_questions = Column(Integer)

class AuditLog(Base):
    __tablename__ = "audit_logs"
    id = Column(Integer, primary_key=True, index=True)
    action = Column(String(100))
    status = Column(String(50))
    details = Column(Text)
    timestamp = Column(DateTime, default=datetime.datetime.utcnow)

def init_db():
    Base.metadata.create_all(bind=engine)
    print("Database tables initialized successfully.")

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

if __name__ == "__main__":
    init_db()
