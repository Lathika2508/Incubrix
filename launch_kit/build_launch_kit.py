import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

kit_dir = r"D:\incubrix project final\launch_kit"
os.makedirs(kit_dir, exist_ok=True)

weeks_data = [
    {
        "week": "Week 1",
        "phase": "Foundation & Technical Fixes",
        "objective": "Deploy high-priority technical SEO and AEO/GEO fixes on incubrix.com",
        "tasks": "Deploy FIX_001 to FIX_005: Organization/SoftwareApplication schema, /llms.txt, FAQPage schema, robots.txt AI bot directives, and HTTPS internal link fixes.",
        "owner": "Growth Engineer & Web Developer",
        "tools": "VS Code, Git, Google Search Console, Rich Results Test",
        "dependencies": "Engineering access to incubrix.com production codebase",
        "budget": "$0 (Internal hours)",
        "deliverables": "Live /llms.txt, valid JSON-LD on homepage/FAQ, updated robots.txt",
        "kpi": "Crawl error rate & Schema validity",
        "proposed_target": "0 validation errors; 100% pass on Google Rich Results Test",
        "risks": "Deployment rollback if staging testing fails",
        "success_criteria": "All 5 critical fixes verified live in production"
    },
    {
        "week": "Week 2",
        "phase": "Foundation & Technical Fixes",
        "objective": "Deploy remaining technical fixes (FIX_006 to FIX_020) and configure indexing monitoring",
        "tasks": "Deploy meta tag optimizations, image alt attributes, Open Graph specifications, XML sitemap modernization, and security.txt.",
        "owner": "Growth Engineer",
        "tools": "Google Search Console, Bing Webmaster Tools, Schema Validator",
        "dependencies": "Week 1 baseline deployment approval",
        "budget": "$0",
        "deliverables": "20/20 technical fixes live and verified",
        "kpi": "Sitemap URL indexation rate",
        "proposed_target": "100% of sitemap URLs submitted and indexed",
        "risks": "Search engine re-crawling latency",
        "success_criteria": "Clean audit report on PageSpeed Insights and zero broken links"
    },
    {
        "week": "Week 3",
        "phase": "Content Publishing Batch 1",
        "objective": "Publish first batch of answer-first pillar pages (PAGE_001 to PAGE_003)",
        "tasks": "Publish 'How to Turn Long Videos into Shorts', 'Best AI Video Generators Comparison', and 'How to Create Faceless Channel'. Configure Article and FAQ schema.",
        "owner": "Content Strategist & SEO Lead",
        "tools": "CMS (Ghost/WordPress/Custom), Grammarly, Schema Markup Generator",
        "dependencies": "Week 1 Schema templates approved",
        "budget": "$150 (Visual assets & infographics)",
        "deliverables": "3 live comprehensive pillar guides with schema and custom diagrams",
        "kpi": "Initial organic impressions & indexing latency",
        "proposed_target": "Indexed within 72 hours of publication",
        "risks": "Thin content penalties if not sufficiently detailed",
        "success_criteria": "All 3 pages index with rich snippet eligibility"
    },
    {
        "week": "Week 4",
        "phase": "Content Publishing Batch 2",
        "objective": "Publish comparison and creator productivity guides (PAGE_004 to PAGE_006)",
        "tasks": "Publish Opus Clip vs Premiere Pro, Caption Writing guide, and Social Automation walkthrough. Link to IncuBrix product features.",
        "owner": "Growth Engineer",
        "tools": "CMS, Loom (demo video embeds)",
        "dependencies": "CMS publishing pipeline active",
        "budget": "$100",
        "deliverables": "3 live comparison and workflow articles",
        "kpi": "Ranking positions for target comparison queries",
        "proposed_target": "First-50 ranking for 'opus clip vs premiere pro'",
        "risks": "Brand comparison pushback",
        "success_criteria": "Objective, fair comparison matrix with clear attribution"
    },
    {
        "week": "Week 5",
        "phase": "Content Publishing Batch 3 & Citation Launch",
        "objective": "Publish remaining articles (PAGE_007 to PAGE_010) and initiate Class A citation outreach",
        "tasks": "Publish content calendar and webinar repurposing guides. Claim official profiles on G2, Capterra, and Trustpilot (C001, C002, C024).",
        "owner": "Partnership Lead & Growth Marketer",
        "tools": "G2 Vendor Portal, Capterra Vendor Portal, Trustpilot",
        "dependencies": "Articles live and internally linked",
        "budget": "$0 (Free tier vendor claiming)",
        "deliverables": "4 published guides (10 total); 3 claimed review directory profiles",
        "kpi": "Directory profile verification",
        "proposed_target": "Verified badges on G2 and Capterra",
        "risks": "Directory approval delays",
        "success_criteria": "Profiles live with accurate pricing and product screenshots"
    },
    {
        "week": "Week 6",
        "phase": "Community & Forum Visibility",
        "objective": "Execute high-value community engagement across Reddit and Quora (C003, C016, C020, C029)",
        "tasks": "Participate in r/NewTubers and r/VideoEditing discussions sharing workflow benchmarks and video retention data without direct promotional links.",
        "owner": "Growth Engineer (Technical Advocate)",
        "tools": "Reddit, Quora, Browser monitoring",
        "dependencies": "Published research guides available for citation reference",
        "budget": "$0",
        "deliverables": "5 insightful community contributions answering creator bottlenecks",
        "kpi": "Upvotes, discussion engagement, and referral traffic",
        "proposed_target": "50+ net upvotes and inclusion in community wiki / discussions SERP carousel",
        "risks": "Moderator removal if viewed as overly promotional",
        "success_criteria": "Strict adherence to subreddit rules; zero account strikes"
    },
    {
        "week": "Week 7",
        "phase": "Weekly Visibility Engine Automation",
        "objective": "Run automated bi-weekly visibility tracker across 100 questions and 4 AI assistants",
        "tasks": "Execute visibility engine refresh. Analyze changes in Google position, snippet presence, and Perplexity/ChatGPT citations.",
        "owner": "Growth Engineer",
        "tools": "IncuBrix Visibility Engine, FastAPI, SQLite, Plotly Dashboard",
        "dependencies": "Initial baseline dataset in SQLite",
        "budget": "$10 (Compute / proxy checks)",
        "deliverables": "Weekly Visibility Progress Report (Markdown, HTML, XLSX)",
        "kpi": "Total tracked visibility score & answer share",
        "proposed_target": "First organic Top 20 ranking detected for long-tail query",
        "risks": "Search engine rate-limiting",
        "success_criteria": "Automated pipeline executes and exports without manual intervention"
    },
    {
        "week": "Week 8",
        "phase": "Authority Backlink & Co-Marketing",
        "objective": "Pitch guest expertise and co-marketing data studies to Tier-2 publications (Zapier, VidIQ, Search Engine Journal)",
        "tasks": "Submit data study on 'AI Captioning Impact on YouTube Shorts Retention' to industry editors (C004, C006, C010).",
        "owner": "PR & Content Strategist",
        "tools": "Hunter.io, Pitchbox, Google Docs",
        "dependencies": "Anonymized retention dataset from IncuBrix product studio",
        "budget": "$200 (Outreach & content drafting)",
        "deliverables": "3 tailored editorial pitches sent to verified editors",
        "kpi": "Editorial response rate",
        "proposed_target": "At least 1 editorial placement or resource link confirmed",
        "risks": "Low editorial response rates",
        "success_criteria": "High-authority editorial backlink (DA >70) secured"
    },
    {
        "week": "Week 9",
        "phase": "Product Hunt Launch & Developer Ecosystem",
        "objective": "Launch IncuBrix on Product Hunt and submit to open-source AI lists (C008, C021)",
        "tasks": "Prepare animated product walkthroughs, maker comments, and submit pull requests to curated GitHub repositories (awesome-generative-ai).",
        "owner": "Full Product & Growth Team",
        "tools": "Product Hunt, GitHub, Twitter / X",
        "dependencies": "Stable app.incubrix.com onboarding flow",
        "budget": "$100",
        "deliverables": "Live Product Hunt launch; 2 merged GitHub directory PRs",
        "kpi": "Product Hunt upvotes and referral signups",
        "proposed_target": "Top 5 Product of the Day; 1,000+ launch day unique visitors",
        "risks": "Server traffic spike handling",
        "success_criteria": "Zero app downtime; 250+ free trial accounts created"
    },
    {
        "week": "Week 10",
        "phase": "Mid-Term Review & Content Brief Scaling",
        "objective": "Analyze top-performing keyword clusters and generate 10 new content briefs using the brief generator",
        "tasks": "Run Content Brief Generator for top 10 rising queries in creator questions dataset. Hand off briefs to freelance writing team.",
        "owner": "Content Strategist",
        "tools": "IncuBrix Content Brief Generator, Google Search Console",
        "dependencies": "Search Console 60-day query impression data",
        "budget": "$500 (Freelance content drafting)",
        "deliverables": "10 production-ready content briefs",
        "kpi": "Content production velocity",
        "proposed_target": "10 new briefs generated and queued for draft",
        "risks": "Quality drift with freelance writers",
        "success_criteria": "All briefs meet strict answer-first editorial guidelines"
    },
    {
        "week": "Week 11",
        "phase": "Conversion Rate Optimization (CRO)",
        "objective": "Optimize blog-to-product conversion funnels and interactive tools",
        "tasks": "Deploy interactive Credit ROI Calculator on /pricing and embedded short video preview widgets on high-traffic blog pages.",
        "owner": "Growth Engineer & UI Designer",
        "tools": "Figma, React/Next.js, PostHog / Google Analytics 4",
        "dependencies": "Traffic volume >5,000 monthly sessions",
        "budget": "$150",
        "deliverables": "Live interactive credit calculator and CTA banners",
        "kpi": "Visitor-to-waitlist conversion rate",
        "proposed_target": "Increase blog conversion rate from 1.5% to 3.8%",
        "risks": "Widget load time affecting Core Web Vitals",
        "success_criteria": "PageSpeed score remains >90 on mobile"
    },
    {
        "week": "Week 12",
        "phase": "GEO & AI Assistant Re-measurement",
        "objective": "Conduct comprehensive AI assistant re-measurement across ChatGPT, Gemini, Perplexity, and Claude",
        "tasks": "Rerun the 100-question baseline. Measure mention frequency, recommendation rate, and third-party citation accuracy.",
        "owner": "Growth Engineer",
        "tools": "IncuBrix Visibility Engine, API connectors",
        "dependencies": "Weeks 1–11 SEO and citation activities completed",
        "budget": "$20",
        "deliverables": "Comparative Visibility Scorecard (Week 1 vs Week 12)",
        "kpi": "AI assistant mention rate",
        "proposed_target": "10%+ AI assistant mention rate for core repurposing queries",
        "risks": "LLM knowledge cutoff latency",
        "success_criteria": "Documented citation of incubrix.com in Perplexity and Google AI Overviews"
    },
    {
        "week": "Week 13",
        "phase": "Executive Synthesis & Next Quarter Roadmap",
        "objective": "Deliver comprehensive 90-day growth retrospective and Q2 scale plan",
        "tasks": "Compile organic traffic, SERP rankings, AI mention rates, and sign-up acquisition costs. Formulate Q2 programmatic SEO expansion.",
        "owner": "Growth Engineer & Head of Growth",
        "tools": "Google Looker Studio, Excel, Streamlit Dashboard",
        "dependencies": "Complete 90-day analytics data",
        "budget": "$0",
        "deliverables": "Executive Q1 Review Presentation & Q2 Programmatic Roadmap",
        "kpi": "Total organic sign-ups and blended Customer Acquisition Cost (CAC)",
        "proposed_target": "1,500+ cumulative organic signups; CAC reduction of 45%",
        "risks": "Market shift towards new video formats",
        "success_criteria": "Sign-off from executive leadership on Q2 expansion budget"
    }
]

# Write XLSX
xlsx_path = os.path.join(kit_dir, "90_day_roadmap.xlsx")
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "90_Day_Growth_Roadmap"

h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
h_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

headers = list(weeks_data[0].keys())
ws.append(headers)
for c_idx in range(1, len(headers)+1):
    c = ws.cell(1, c_idx)
    c.fill = h_fill
    c.font = h_font
    c.alignment = Alignment(horizontal="center", vertical="center")

for row in weeks_data:
    ws.append([row[k] for k in headers])

for col in ws.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = openpyxl.utils.get_column_letter(col[0].column)
    ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

wb.save(xlsx_path)
print(f"Saved {xlsx_path}")

# Write Markdown Roadmap
md_path = os.path.join(kit_dir, "90_day_roadmap.md")
with open(md_path, "w", encoding="utf-8") as f_md:
    f_md.write("""# IncuBrix 90-Day Growth Engineering Roadmap (13 Weeks)

## Strategic Objective
Transform IncuBrix from an unranked baseline into the leading organic search authority and AI assistant recommendation for video creators seeking AI video repurposing, dynamic captioning, and multi-platform publishing.

---

## 13-Week Operational Execution Plan

""")
    for w in weeks_data:
        f_md.write(f"""### {w['week']}: {w['objective']}
- **Phase:** {w['phase']}
- **Owner:** `{w['owner']}` | **Budget:** `{w['budget']}`
- **Core Tasks:** {w['tasks']}
- **Tools Used:** {w['tools']}
- **Dependencies:** {w['dependencies']}
- **Key Deliverables:** {w['deliverables']}
- **Target KPI:** {w['kpi']} -> **Target:** `{w['proposed_target']}`
- **Risk Mitigation:** {w['risks']}
- **Success Criteria:** {w['success_criteria']}

---
""")

print(f"Saved {md_path}")

# Write Launch Plan document
plan_path = os.path.join(kit_dir, "launch_plan.md")
with open(plan_path, "w", encoding="utf-8") as f_plan:
    f_plan.write("""# IncuBrix Growth & Visibility Engine Launch Plan

## 1. Day 1 Operational Launch (Next Monday Checklist)
This plan is 100% turnkey and executable next Monday with zero rework:

1. **Morning (09:00 - 12:00): Technical SEO Deployment**
   - Apply `FIX_001` (Organization & SoftwareApplication schema) and `FIX_003` (FAQPage schema) to `<head>` templates.
   - Deploy `FIX_002` (`/llms.txt`) and `FIX_004` (`robots.txt` AI user-agent rules).
   - Fix internal HTTP links to HTTPS (`FIX_005`).
   - Run verification in Google Rich Results Test.

2. **Afternoon (13:00 - 15:00): Publishing Setup**
   - Stage the first 3 publish-ready articles (`PAGE_001`, `PAGE_002`, `PAGE_003`) in the CMS.
   - Verify canonical tags, meta titles, descriptions, and Open Graph previews.
   - Submit updated `sitemap.xml` in Google Search Console.

3. **Evening (15:00 - 17:00): Visibility Engine & Review Directory Setup**
   - Initialize the local Visibility Engine database (`python -m backend.app.database`).
   - Claim vendor profiles on G2, Capterra, and Trustpilot (`C001`, `C002`, `C024`).
   - Launch Streamlit reporting dashboard (`streamlit run streamlit_app.py`).

---

## 2. Resource & Budget Allocation
| Resource Category | Tool / Vendor | Estimated Monthly Budget | Responsibility |
| :--- | :--- | :--- | :--- |
| **SEO & Crawling** | Google Search Console, Free Crawlers | $0 | Growth Engineer |
| **Directory Profiles** | G2, Capterra, Product Hunt | $0 (Free maker tiers) | Partnership Lead |
| **Creative Assets** | Canva / Figma | $30 / month | Content Designer |
| **Content Freelancers** | Specialized creator writers | $500 / month (Weeks 6–12) | Content Strategist |
| **Total Monthly Burn** | | **~$530 / month** | |

---

## 3. 30-Day and 90-Day Targets
| Benchmark Metric | Current Baseline (Day 1) | 30-Day Target | 90-Day Target |
| :--- | :--- | :--- | :--- |
| **Technical Schema Errors** | 100% missing | **0 errors (100% valid)** | 0 errors |
| **Publish-Ready Pages Live** | 0 live | **10 live articles** | 25+ live articles |
| **Top 20 Google Rankings** | 0 queries | **5 queries** | 35+ queries |
| **Featured Snippet Placements** | 0 snippets | **2 placements** | 8+ placements |
| **Perplexity / Claude Mentions**| 0% | **5% mention rate** | 18% mention rate |
| **Monthly Organic Signups** | Baseline unmeasured | **150 signups** | 1,500+ signups |

---

## 4. Risk Matrix & Handling Strategy
1. **Search Engine Indexation Lag:** Mitigate by submitting direct URL inspect requests in Search Console and leveraging IndexNow API.
2. **LLM Knowledge Cutoff Delay:** Accelerate retrieval-augmented generation (RAG) by obtaining verified citations on actively indexed web properties (Zapier, G2, Reddit).
3. **Engineering Resource Constraints:** All 20 technical fixes include ready-to-paste code snippets requiring less than 2 hours of front-end engineering time.
""")

print(f"Saved {plan_path}")
print("Successfully generated all launch kit documents.")
