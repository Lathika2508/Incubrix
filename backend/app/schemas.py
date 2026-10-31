from pydantic import BaseModel
from typing import List, Optional
import datetime

class QuestionBase(BaseModel):
    question_id: str
    question: str
    category: str
    intent: str
    source_type: str
    source_url: str
    market: str
    priority: str = "Medium"

class QuestionCreate(QuestionBase):
    pass

class QuestionResponse(QuestionBase):
    id: int
    created_at: Optional[datetime.datetime] = None

    class Config:
        from_attributes = True

class MeasurementBase(BaseModel):
    question_id: str
    day: str
    measurement_date: str
    google_position: str
    answer_box_present: bool = False
    quoted_source_url: Optional[str] = None
    chatgpt_mentioned: bool = False
    gemini_mentioned: bool = False
    perplexity_mentioned: bool = False
    claude_copilot_mentioned: bool = False
    competitors_named: Optional[str] = None
    raw_response_path: Optional[str] = None
    changed_from_previous: bool = False

class MeasurementResponse(MeasurementBase):
    id: int

    class Config:
        from_attributes = True

class TechnicalFixBase(BaseModel):
    fix_id: str
    page_url: str
    problem: str
    evidence_url: str
    fix_summary: str
    code_path: str
    validator_used: str
    validated: str = "Yes"
    impact: str = "Medium"
    effort: str = "Low"

class ContentPageBase(BaseModel):
    item_id: str
    slug: str
    title: str
    primary_query: str
    target_qids: str
    draft_path: str
    status: str = "Ready"

class CitationTargetBase(BaseModel):
    target_id: str
    website: str
    target_url: str
    page_type: str
    classification: str
    priority: str
    mapped_qids: str

class ContentBriefRequest(BaseModel):
    creator_question: str
    category: str
    search_intent: str
    market: str = "US"
    primary_keyword: str
    secondary_keywords: List[str] = []
    competitors: List[str] = []

class ContentBriefResponse(BaseModel):
    seo_title: str
    meta_description: str
    answer_first_intro: str
    recommended_slug: str
    article_structure: List[str]
    faqs: List[dict]
    internal_links: List[str]
    structured_data_type: str
    quality_checklist: List[str]
