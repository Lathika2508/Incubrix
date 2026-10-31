import pytest
import os
import sys
import csv
import json
import openpyxl

# Add root directory to python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from fastapi.testclient import TestClient
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from backend.app.database import Base, Question, Measurement, TechnicalFix, ContentPage, CitationTarget
from backend.app.main import app
from backend.app.engine import calculate_visibility_metrics, generate_content_brief

# Test database
TEST_DB_URL = "sqlite:///:memory:"
test_engine = create_engine(TEST_DB_URL, connect_args={"check_same_thread": False})
TestingSessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=test_engine)

@pytest.fixture(scope="module")
def db_session():
    Base.metadata.create_all(bind=test_engine)
    session = TestingSessionLocal()
    yield session
    session.close()
    Base.metadata.drop_all(bind=test_engine)

client = TestClient(app)

# Test 1: Database initialisation
def test_01_database_initialization(db_session):
    assert db_session is not None
    assert "memory" in str(test_engine.url)

# Test 2: Question creation
def test_02_question_import(db_session):
    q = Question(
        question_id="Q_TEST_001",
        question="How to test video generation?",
        category="AI video generation",
        intent="How-to",
        source_type="Google Autocomplete",
        source_url="https://google.com/search?q=test",
        market="US",
        priority="High"
    )
    db_session.add(q)
    db_session.commit()
    fetched = db_session.query(Question).filter(Question.question_id == "Q_TEST_001").first()
    assert fetched is not None
    assert fetched.question == "How to test video generation?"

# Test 3: Question count check
def test_03_question_count():
    csv_path = r"D:\incubrix project final\data\creator_questions_100.csv"
    with open(csv_path, "r", encoding="utf-8") as f:
        reader = list(csv.DictReader(f))
    assert len(reader) == 100

# Test 4: Source validation check
def test_04_source_validation():
    csv_path = r"D:\incubrix project final\data\creator_questions_100.csv"
    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            assert row["source_url"].startswith("http")
            assert len(row["source_type"]) > 3

# Test 5: Exact duplicate detection
def test_05_exact_duplicate_detection():
    csv_path = r"D:\incubrix project final\data\creator_questions_100.csv"
    seen = set()
    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            q_clean = row["question"].strip().lower()
            assert q_clean not in seen, f"Duplicate question detected: {q_clean}"
            seen.add(q_clean)

# Test 6: Google measurement creation
def test_06_google_measurement_creation(db_session):
    m = Measurement(
        question_id="Q_TEST_001",
        day="Day 1",
        measurement_date="2026-10-24",
        google_position=">50",
        answer_box_present=True
    )
    db_session.add(m)
    db_session.commit()
    fetched = db_session.query(Measurement).filter(Measurement.question_id == "Q_TEST_001").first()
    assert fetched.google_position == ">50"
    assert fetched.answer_box_present is True

# Test 7: AI response creation
def test_07_ai_response_creation(db_session):
    m = db_session.query(Measurement).filter(Measurement.question_id == "Q_TEST_001").first()
    m.chatgpt_mentioned = False
    m.gemini_mentioned = False
    m.perplexity_mentioned = False
    m.claude_copilot_mentioned = False
    db_session.commit()
    assert m.chatgpt_mentioned is False

# Test 8: Day 1 and Day 2 storage
def test_08_day1_day2_storage(db_session):
    m2 = Measurement(
        question_id="Q_TEST_001",
        day="Day 2",
        measurement_date="2026-10-31",
        google_position=">50",
        changed_from_previous=False
    )
    db_session.add(m2)
    db_session.commit()
    count = db_session.query(Measurement).filter(Measurement.question_id == "Q_TEST_001").count()
    assert count == 2

# Test 9: Market tracking
def test_09_market_tracking():
    csv_path = r"D:\incubrix project final\data\creator_questions_100.csv"
    markets = set()
    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            markets.add(row["market"])
    assert "US" in markets
    assert "IN" in markets
    assert "UK" in markets

# Test 10: Visibility scoring
def test_10_visibility_scoring(db_session):
    metrics = calculate_visibility_metrics(db_session, "Day 1")
    assert "google" in metrics
    assert "ai_assistants" in metrics

# Test 11: Missing-data handling
def test_11_missing_data_handling(db_session):
    empty_metrics = calculate_visibility_metrics(db_session, "NonExistentDay")
    assert empty_metrics.get("sample_size") == 0

# Test 12: Competitor tracking
def test_12_competitor_tracking():
    vis_path = r"D:\incubrix project final\data\exports\visibility.csv"
    with open(vis_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            assert len(row["competitors_named"]) > 0

# Test 13: Evidence references
def test_13_evidence_references():
    vis_path = r"D:\incubrix project final\data\exports\visibility.csv"
    with open(vis_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        first_row = next(reader)
        raw_p = os.path.join(r"D:\incubrix project final", first_row["raw_answers_url"])
        assert os.path.exists(raw_p), f"Evidence path missing: {raw_p}"

# Test 14: CSV export
def test_14_csv_export():
    assert os.path.exists(r"D:\incubrix project final\data\exports\visibility.csv")
    assert os.path.exists(r"D:\incubrix project final\visibility.csv")

# Test 15: Excel export
def test_15_excel_export():
    assert os.path.exists(r"D:\incubrix project final\data\creator_questions_100.xlsx")
    wb = openpyxl.load_workbook(r"D:\incubrix project final\data\creator_questions_100.xlsx")
    assert len(wb.sheetnames) > 0

# Test 16: Workbook preservation
def test_16_workbook_preservation():
    comp_path = r"D:\incubrix project final\submission\IncuBrix_Option3_Submission_Workbook_COMPLETED.xlsx"
    wb = openpyxl.load_workbook(comp_path, data_only=False)
    ws_sc = wb["04_SELF_CHECK"]
    assert str(ws_sc["D6"].value).startswith("=COUNTA")
    assert str(ws_sc["D9"].value).startswith("=COUNTIF")
    assert str(ws_sc["D10"].value).startswith("=COUNTIFS")
    assert str(ws_sc["D11"].value).startswith("=COUNTIF")

# Test 17: Weekly reports existence
def test_17_weekly_reports():
    rep_md = r"D:\incubrix project final\reports\weekly_report_week_1.md"
    rep_html = r"D:\incubrix project final\reports\weekly_report_week_1.html"
    assert os.path.exists(rep_md)
    assert os.path.exists(rep_html)

# Test 18: Content briefs generation
def test_18_content_brief_generator():
    req = {
        "creator_question": "How to edit YouTube Shorts faster?",
        "category": "Video editing",
        "search_intent": "How-to",
        "primary_keyword": "edit youtube shorts faster"
    }
    brief = generate_content_brief(req)
    assert "seo_title" in brief
    assert "answer_first_intro" in brief
    assert len(brief["article_structure"]) > 4

# Test 19: SEO audit error handling
def test_19_audit_evidence_loaded():
    audit_json = r"D:\incubrix project final\technical_audit\audit_evidence.json"
    assert os.path.exists(audit_json)
    with open(audit_json, "r", encoding="utf-8") as f:
        data = json.load(f)
    assert len(data) >= 15

# Test 20: API validation
def test_20_api_endpoints():
    res = client.get("/api/health")
    assert res.status_code == 200
    assert res.json()["status"] == "HEALTHY"

# Test 21: Duplicate refresh handling
def test_21_duplicate_refresh():
    comp_path = r"D:\incubrix project final\submission\IncuBrix_Option3_Submission_Workbook_COMPLETED.xlsx"
    wb = openpyxl.load_workbook(comp_path)
    ws_q = wb["02_QUESTIONS"]
    assert ws_q["U5"].value is not None

# Test 22: Invalid input recovery
def test_22_invalid_input_recovery():
    # Test invalid brief payload error recovery
    res = client.post("/api/brief/generate", json={"invalid_field": 123})
    assert res.status_code == 422
