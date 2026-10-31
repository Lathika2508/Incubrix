import json
import urllib.parse
import os
import datetime
import requests

seeds_by_market = [
    # US Market Seeds
    ("US", "how to create videos using ai", "AI video generation", "How-to"),
    ("US", "how to turn long videos into shorts", "Repurposing long videos into short videos", "How-to"),
    ("US", "how to repurpose podcast into social media clips", "Content workflows", "How-to"),
    ("US", "best ai video generator for youtube shorts", "YouTube Shorts", "Comparison"),
    ("US", "how to make faceless youtube channel with ai", "AI content creation", "How-to"),
    ("US", "how to edit youtube shorts faster", "Video editing", "Problem-solving"),
    ("US", "how to create viral tiktok videos using ai", "TikTok content", "How-to"),
    ("US", "opus clip vs premiere pro", "Comparison queries", "Comparison"),
    ("US", "best ai tools for content creators 2026", "AI content creation", "Commercial investigation"),
    ("US", "how to automate social media posting with ai", "Social media automation", "How-to"),
    ("US", "how to plan 30 days of content in one day", "Content planning", "How-to"),
    ("US", "how to write engaging hooks for reels", "Instagram Reels", "How-to"),
    ("US", "how to get more views on youtube shorts algorithm", "YouTube channel growth", "Informational"),
    ("US", "best ai tool to repurpose webinar into clips", "Repurposing long videos into short videos", "Commercial investigation"),
    ("US", "how to add dynamic captions to reels automatically", "Video editing", "How-to"),
    ("US", "how to grow on instagram using ai tools", "Instagram growth", "How-to"),
    ("US", "best ai voiceover generator for creators", "AI image and video tools", "Comparison"),
    ("US", "how to repurpose youtube video for linkedin", "Content workflows", "How-to"),

    # IN Market Seeds
    ("IN", "how to write instagram captions with ai free", "Social media captions", "How-to"),
    ("IN", "how to schedule posts on instagram for free", "Publishing and scheduling", "How-to"),
    ("IN", "how to grow a youtube channel with ai", "YouTube channel growth", "How-to"),
    ("IN", "capcut vs premiere pro for reels", "Comparison queries", "Comparison"),
    ("IN", "best free video editing app for youtube shorts", "Video editing", "Comparison"),
    ("IN", "how to make trending instagram reels with ai", "Instagram Reels", "How-to"),
    ("IN", "how to add auto captions to video free", "Video editing", "How-to"),
    ("IN", "how to get content ideas for youtube channel", "Content ideas", "Problem-solving"),
    ("IN", "best social media scheduler for creators", "Publishing and scheduling", "Commercial investigation"),
    ("IN", "how to earn money from faceless youtube channel", "YouTube channel growth", "Informational"),
    ("IN", "how to generate script for youtube video using ai", "AI writing tools", "How-to"),
    ("IN", "canva vs capcut for instagram reels", "Comparison queries", "Comparison"),
    ("IN", "how to create bulk reels in canva with chatgpt", "Content workflows", "How-to"),
    ("IN", "how to increase instagram engagement rate organically", "Audience engagement", "Problem-solving"),

    # UK Market Seeds
    ("UK", "how to plan a month of social media content", "Content planning", "How-to"),
    ("UK", "best ai tools for content creators uk", "AI content creation", "Commercial investigation"),
    ("UK", "how to automate social media content creation", "Social media automation", "How-to"),
    ("UK", "how to batch produce content for instagram", "Creator productivity", "How-to"),
    ("UK", "buffer vs hootsuite for small creators", "Comparison queries", "Comparison"),
    ("UK", "how to build a content calendar for youtube", "Content calendars", "How-to"),
    ("UK", "how to repurpose long form content into newsletters and shorts", "Content workflows", "How-to"),
    ("UK", "best ai caption generator for tiktok", "Social media captions", "Commercial investigation"),
    ("UK", "how to measure social media content performance", "Content marketing", "Informational"),
    ("UK", "how to create b-roll footage with ai", "AI image and video tools", "How-to"),
    ("UK", "how to repurpose zoom recordings into short clips", "Repurposing long videos into short videos", "How-to"),
    ("UK", "how to maintain consistent posting schedule across platforms", "Publishing and scheduling", "Problem-solving")
]

all_queries = []
seen_questions = set()

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
}

now = datetime.datetime.now()
timestamp_str = now.strftime("%Y-%m-%d %H:%M:%S")

print(f"Starting discovery at {timestamp_str}...")

for market, seed, category, default_intent in seeds_by_market:
    gl = market.lower()
    encoded = urllib.parse.quote(seed)
    url = f"https://suggestqueries.google.com/complete/search?client=chrome&hl=en&gl={gl}&q={encoded}"
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            suggestions = data[1]
            # include the seed itself if framed as question or specific intent
            candidates = [seed] + suggestions
            for c in candidates:
                q_clean = c.strip().lower()
                # filter out very short non-queries
                if len(q_clean) < 15:
                    continue
                if q_clean not in seen_questions:
                    seen_questions.add(q_clean)
                    # Determine intent
                    intent = default_intent
                    if "vs" in q_clean or " or " in q_clean:
                        intent = "Comparison"
                    elif q_clean.startswith("how to") or q_clean.startswith("how do"):
                        intent = "How-to"
                    elif q_clean.startswith("best") or "top" in q_clean or "review" in q_clean:
                        intent = "Commercial investigation"
                    elif "why" in q_clean or "what is" in q_clean or "how does" in q_clean:
                        intent = "Informational"

                    all_queries.append({
                        "question": c.strip(),
                        "category": category,
                        "intent": intent,
                        "source_type": "Google Autocomplete",
                        "source_url": f"https://www.google.com/search?q={urllib.parse.quote(c.strip())}&gl={gl}",
                        "source_platform": "Google Search / Suggest",
                        "market": market,
                        "search_location": f"Google {market} (gl={gl}, hl=en)",
                        "collection_date": "2026-10-31",
                        "collection_time": now.strftime("%H:%M:%S"),
                        "seed_used": seed
                    })
    except Exception as e:
        print(f"Error querying {seed}: {e}")

print(f"Total discovered unique questions: {len(all_queries)}")

# Save raw discovery
os.makedirs(r"D:\incubrix project final\data\raw", exist_ok=True)
with open(r"D:\incubrix project final\data\raw\google_suggest_raw.json", "w", encoding="utf-8") as f:
    json.dump(all_queries, f, indent=2)

print("Saved raw discovery to data/raw/google_suggest_raw.json")
