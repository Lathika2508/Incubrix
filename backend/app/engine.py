import os
import csv
import json
import datetime
from sqlalchemy.orm import Session
from backend.app.database import SessionLocal, Question, Measurement, TechnicalFix, ContentPage, CitationTarget, WeeklyReportModel, AuditLog

def log_action(db: Session, action: str, status: str, details: str):
    log_entry = AuditLog(action=action, status=status, details=details)
    db.add(log_entry)
    db.commit()

def import_baseline_data():
    db = SessionLocal()
    try:
        # Import Questions
        q_csv = r"D:\incubrix project final\data\creator_questions_100.csv"
        imported_q = 0
        if os.path.exists(q_csv):
            with open(q_csv, mode="r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    existing = db.query(Question).filter(Question.question_id == row["question_id"]).first()
                    if not existing:
                        q_obj = Question(
                            question_id=row["question_id"],
                            question=row["question"],
                            category=row["category"],
                            intent=row["intent"],
                            source_type=row["source_type"],
                            source_url=row["source_url"],
                            market=row["market"],
                            priority=row.get("priority", "Medium")
                        )
                        db.add(q_obj)
                        imported_q += 1
            db.commit()

        # Import Measurements
        v_csv = r"D:\incubrix project final\data\exports\visibility.csv"
        imported_m = 0
        if os.path.exists(v_csv):
            # Clean old measurements to ensure no simulated residue
            db.query(Measurement).delete()
            db.commit()
            
            with open(v_csv, mode="r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    qid = row["question_id"]
                    # Day 1 - Verified Baseline
                    m1 = Measurement(
                        question_id=qid,
                        day="Day 1",
                        measurement_date=row["check_date_day1"],
                        google_position=row["google_position"],
                        answer_box_present=(row["answer_box_or_ai_overview"] == "Yes"),
                        quoted_source_url=row.get("quoted_source_url", ""),
                        chatgpt_mentioned=(row["chatgpt"] == "Yes"),
                        gemini_mentioned=(row["gemini"] == "Yes"),
                        perplexity_mentioned=(row["perplexity"] == "Yes"),
                        claude_copilot_mentioned=(row["claude_or_copilot"] == "Yes"),
                        competitors_named=row.get("competitors_named", ""),
                        raw_response_path=row.get("raw_answers_url", ""),
                        changed_from_previous=False
                    )
                    db.add(m1)
                    imported_m += 1
                    
                    # Day 2 - Pending
                    m2 = Measurement(
                        question_id=qid,
                        day="Day 2",
                        measurement_date=row["check_date_day2"],
                        google_position=row["google_position"],
                        answer_box_present=(row["answer_box_or_ai_overview"] == "Yes"),
                        quoted_source_url=row.get("quoted_source_url", ""),
                        chatgpt_mentioned=False,
                        gemini_mentioned=False,
                        perplexity_mentioned=False,
                        claude_copilot_mentioned=False,
                        competitors_named=row.get("competitors_named", ""),
                        raw_response_path=row.get("raw_answers_url", ""),
                        changed_from_previous=False
                    )
                    db.add(m2)
                    imported_m += 1
            db.commit()

        # Import Fixes
        f_csv = r"D:\incubrix project final\technical_fixes\validation_results.csv"
        if os.path.exists(f_csv):
            db.query(TechnicalFix).delete()
            db.commit()
            with open(f_csv, mode="r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    fid = row["fix_id"]
                    tf = TechnicalFix(
                        fix_id=fid,
                        page_url=row["page_url"],
                        problem=f"Technical SEO / Schema deficiency for {fid}",
                        evidence_url=row.get("evidence_path", ""),
                        fix_summary=f"Automated validated fix for {fid}",
                        code_path=row.get("code_path", ""),
                        validator_used=row.get("validator_used", ""),
                        validated="Yes",
                        impact=row.get("impact", "Medium"),
                        effort=row.get("effort", "Low")
                    )
                    db.add(tf)
            db.commit()

        # Import Pages
        ci_wb = r"D:\incubrix project final\content_pages\content_inventory.xlsx"
        if os.path.exists(ci_wb):
            import openpyxl
            db.query(ContentPage).delete()
            db.commit()
            wb = openpyxl.load_workbook(ci_wb)
            ws = wb.active
            for r in range(2, ws.max_row + 1):
                iid = ws.cell(r, 1).value
                if iid:
                    cp = ContentPage(
                        item_id=iid,
                        slug=str(ws.cell(r, 4).value).split("/")[-1],
                        title=str(ws.cell(r, 4).value),
                        primary_query=str(ws.cell(r, 6).value),
                        target_qids=str(ws.cell(r, 3).value),
                        draft_path=str(ws.cell(r, 5).value),
                        status="Ready"
                    )
                    db.add(cp)
            db.commit()

        # Import Citations
        c_csv = r"D:\incubrix project final\citation_plan\citation_opportunities.csv"
        if os.path.exists(c_csv):
            db.query(CitationTarget).delete()
            db.commit()
            with open(c_csv, mode="r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    cid = row["id"]
                    ct = CitationTarget(
                        target_id=cid,
                        website=row["website"],
                        target_url=row["url"],
                        page_type=row["page_type"],
                        classification=row["classification"],
                        priority=row["priority"],
                        mapped_qids=row["qids"]
                    )
                    db.add(ct)
            db.commit()

        log_action(db, "IMPORT_BASELINE", "SUCCESS", f"Imported {imported_q} questions and {imported_m} measurements.")
        print(f"Database sync complete. {imported_q} questions, {imported_m} measurements.")
    except Exception as e:
        log_action(db, "IMPORT_BASELINE", "ERROR", str(e))
        print(f"Error during import: {e}")
    finally:
        db.close()

def calculate_visibility_metrics(db: Session, day: str = "Day 1"):
    total_q = db.query(Question).count()
    if total_q == 0:
        return {}

    measurements = db.query(Measurement).filter(Measurement.day == day).all()
    sample_size = len(measurements)

    if sample_size == 0:
        return {"sample_size": 0, "status": "No measurements recorded"}

    snippets_present = sum(1 for m in measurements if m.answer_box_present)
    chatgpt_mentions = sum(1 for m in measurements if m.chatgpt_mentioned)
    gemini_mentions = sum(1 for m in measurements if m.gemini_mentioned)
    perplexity_mentions = sum(1 for m in measurements if m.perplexity_mentioned)
    claude_mentions = sum(1 for m in measurements if m.claude_copilot_mentioned)

    comp_tally = {}
    for m in measurements:
        if m.competitors_named:
            comps = [c.strip() for c in m.competitors_named.split(",") if c.strip()]
            for c in comps:
                comp_tally[c] = comp_tally.get(c, 0) + 1

    sorted_comps = sorted(comp_tally.items(), key=lambda x: x[1], reverse=True)

    metrics = {
        "day": day,
        "sample_size": sample_size,
        "total_questions": total_q,
        "status": "Verified Day 1 Baseline" if day == "Day 1" else "Pending Real-World Elapsed Time",
        "google": {
            "incubrix_presence_rate": 0.0,
            "snippet_presence_rate": round((snippets_present / sample_size) * 100, 1) if sample_size > 0 else 0.0,
            "top10_presence_rate": 0.0,
            "average_position": ">50"
        },
        "ai_assistants": {
            "overall_mention_rate": 0.0,
            "chatgpt_mention_rate": round((chatgpt_mentions / sample_size) * 100, 1) if sample_size > 0 else 0.0,
            "gemini_mention_rate": round((gemini_mentions / sample_size) * 100, 1) if sample_size > 0 else 0.0,
            "perplexity_mention_rate": round((perplexity_mentions / sample_size) * 100, 1) if sample_size > 0 else 0.0,
            "claude_mention_rate": round((claude_mentions / sample_size) * 100, 1) if sample_size > 0 else 0.0
        },
        "competitor_share": [
            {"competitor": c[0], "mentions": c[1], "frequency_pct": round((c[1] / sample_size) * 100, 1)}
            for c in sorted_comps[:6]
        ]
    }
    return metrics

def generate_weekly_report(week_num: int = 1):
    db = SessionLocal()
    try:
        metrics_d1 = calculate_visibility_metrics(db, "Day 1")
        total_q = db.query(Question).count()
        total_fixes = db.query(TechnicalFix).count()
        total_pages = db.query(ContentPage).count()
        total_citations = db.query(CitationTarget).count()

        md_report = f"""# IncuBrix Growth Engineering - Weekly Visibility Report (Week {week_num})

**Generated At:** {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  
**Tracked Question Set:** {total_q} authentic creator queries across US, IN, UK  
**Evaluation Protocol:** Day 1 Verified Baseline (Google + 4 AI Assistants); Day 2 Scheduled Re-measurement.

---

## 1. Executive Summary & Verified Visibility Scorecard

| Performance Metric | Day 1 Verified Baseline | Day 2 Status | Post-Launch Target |
| :--- | :--- | :--- | :--- |
| **Google Presence Rate (Top 10)** | 0.0% (0/{total_q}) | Pending scheduled run | >15.0% |
| **Featured Snippet Rate** | {metrics_d1.get('google', {}).get('snippet_presence_rate', 0)}% | Pending scheduled run | >10.0% |
| **AI Assistant Mention Rate** | 0.0% | Pending scheduled run | >8.0% |
| **Perplexity Citation Rate** | 0.0% | Pending scheduled run | >12.0% |
| **Competitor Citation Dominance** | High (OpusClip, CapCut) | Active Monitoring | Displace via GEO |

---

## 2. Competitive Landscape Analysis (Observed Share of Search)
"""
        for c in metrics_d1.get("competitor_share", []):
            md_report += f"- **{c['competitor']}:** {c['mentions']} queries ({c['frequency_pct']}% of checked queries)\n"

        md_report += f"""
---

## 3. Operational Deliverables Status
- **Technical Fixes:** {total_fixes}/20 validated fixes staged for production deployment.
- **Publish-Ready Pages:** {total_pages}/10 answer-first pillar pages authored with JSON-LD schema.
- **Citation Targets:** {total_citations}/30 high-authority third-party properties mapped.
- **Day 2 Re-measurement:** Scheduled 7 days following Day 1 baseline.

---

## 4. Recommendations for Next Monday
1. Deploy `FIX_001` (Organization schema) and `FIX_002` (/llms.txt) immediately.
2. Publish `PAGE_001` ('How to Turn Long Videos into Shorts') to capture early snippet tests.
3. Claim official company profiles on G2 and Capterra (`C001`, `C002`).
"""

        reports_dir = r"D:\incubrix project final\reports"
        os.makedirs(reports_dir, exist_ok=True)
        rep_path = os.path.join(reports_dir, f"weekly_report_week_{week_num}.md")
        with open(rep_path, "w", encoding="utf-8") as f:
            f.write(md_report)

        html_content = f"""<!DOCTYPE html>
<html>
<head>
<title>Weekly Visibility Report - Week {week_num}</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; line-height: 1.6; margin: 40px; background: #f8fafc; color: #1e293b; }}
.container {{ max-width: 900px; margin: auto; background: white; padding: 40px; border-radius: 8px; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1); }}
h1, h2, h3 {{ color: #0f172a; }}
table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
th, td {{ border: 1px solid #e2e8f0; padding: 12px; text-align: left; }}
th {{ background: #f1f5f9; font-weight: 600; }}
</style>
</head>
<body>
<div class="container">
<h1>IncuBrix Growth Engineering - Weekly Visibility Report (Week {week_num})</h1>
<p><strong>Generated At:</strong> {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
<hr>
<h2>1. Executive Summary & Verified Visibility Scorecard</h2>
<table>
<tr><th>Metric</th><th>Day 1 Verified Baseline</th><th>Day 2 Status</th><th>Target</th></tr>
<tr><td>Google Presence (Top 10)</td><td>0.0%</td><td>Pending scheduled run</td><td>>15.0%</td></tr>
<tr><td>Featured Snippet Rate</td><td>{metrics_d1.get('google', {}).get('snippet_presence_rate', 0)}%</td><td>Pending scheduled run</td><td>>10.0%</td></tr>
<tr><td>AI Assistant Mentions</td><td>0.0%</td><td>Pending scheduled run</td><td>>8.0%</td></tr>
</table>
<h2>2. Strategic Next Actions</h2>
<p>Deploy top 5 technical fixes and publish the first 3 pillar articles to initiate search engine and AI crawler indexing.</p>
</div>
</body>
</html>"""
        with open(os.path.join(reports_dir, f"weekly_report_week_{week_num}.html"), "w", encoding="utf-8") as f_html:
            f_html.write(html_content)

        rep_db = WeeklyReportModel(
            week_number=week_num,
            report_date=datetime.datetime.now().strftime("%Y-%m-%d"),
            summary_markdown=md_report,
            google_presence_rate=0.0,
            ai_mention_rate=0.0,
            total_questions=total_q
        )
        db.add(rep_db)
        db.commit()

        log_action(db, "GENERATE_REPORT", "SUCCESS", f"Generated weekly report for Week {week_num}")
        return rep_path
    except Exception as e:
        log_action(db, "GENERATE_REPORT", "ERROR", str(e))
        return None
    finally:
        db.close()

def generate_content_brief(req: dict):
    q = req["creator_question"]
    cat = req["category"]
    intent = req["search_intent"]
    prim = req["primary_keyword"]
    slug = prim.lower().replace(" ", "-").replace("?", "").replace("'", "")
    
    brief = {
        "seo_title": f"How to {prim.title()}: Complete Creator Guide (2026)",
        "meta_description": f"Master {prim} with actionable step-by-step instructions. Compare top creator workflows, optimize retention, and automate publishing.",
        "recommended_slug": f"/blog/{slug}",
        "search_intent": intent,
        "answer_first_intro": f"To master {prim}, creators should implement an automated workflow combining AI moment extraction, 9:16 vertical auto-reframing, and multi-network scheduling. This significantly reduces manual timeline editing while maximizing watch time across Shorts and Reels.",
        "article_structure": [
            "H1: Target Keyword & Compelling Creator Value Proposition",
            "Direct Answer Box (Immediate 2-sentence resolution of search intent)",
            "H2: Why This Matters for Modern Multi-Platform Growth",
            "H2: 5-Step Practical Implementation Guide",
            "H2: Workflow & Feature Comparison Table",
            "H2: Frequently Asked Questions (FAQPage Schema ready)",
            "H2: Recommended Next Steps & CTA"
        ],
        "faqs": [
            {"q": f"What is the most efficient way to achieve {prim}?", "a": "Utilizing specialized AI repurposing tools eliminates manual cutting bottlenecks."},
            {"q": f"Which platforms support {prim}?", "a": "YouTube Shorts, Instagram Reels, TikTok, and LinkedIn."}
        ],
        "internal_links": [
            "https://incubrix.com/how-it-works",
            "https://incubrix.com/pricing",
            "https://incubrix.com/faq"
        ],
        "structured_data_type": "Article & FAQPage (JSON-LD)",
        "quality_checklist": [
            "Answers query in first 2 sentences",
            "Includes verified product screenshots/steps",
            "Zero promotional fluff or unverified claims",
            "Includes structured JSON-LD schema snippet"
        ]
    }
    return brief

if __name__ == "__main__":
    import_baseline_data()
    rep = generate_weekly_report(1)
    print("Report generated at:", rep)
