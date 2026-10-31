import requests
from bs4 import BeautifulSoup
import json
import csv
import os
import datetime

audit_urls = [
    "https://incubrix.com/",
    "https://incubrix.com/about",
    "https://incubrix.com/how-it-works",
    "https://incubrix.com/pricing",
    "https://incubrix.com/contact",
    "https://incubrix.com/waitlist",
    "https://incubrix.com/faq",
    "https://incubrix.com/team",
    "https://incubrix.com/privacy",
    "https://incubrix.com/terms",
    "https://incubrix.com/manifesto",
    "https://incubrix.com/careers",
    "https://incubrix.com/robots.txt",
    "https://incubrix.com/sitemap.xml",
    "https://incubrix.com/llms.txt"
]

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36 IncuBrixAuditBot/1.0"
}

results = []
timestamp_now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

print(f"Starting Technical Audit of incubrix.com at {timestamp_now}...")

for url in audit_urls:
    start_time = datetime.datetime.now()
    try:
        r = requests.get(url, headers=headers, timeout=12)
        latency_ms = int((datetime.datetime.now() - start_time).total_seconds() * 1000)
        status_code = r.status_code
        content_type = r.headers.get("Content-Type", "")
        content_len = len(r.content)
        
        title = ""
        meta_desc = ""
        canonical = ""
        robots_meta = ""
        h1_list = []
        h2_list = []
        og_title = ""
        og_image = ""
        schema_types = []
        images_missing_alt = 0
        total_images = 0
        viewport_found = False
        
        if "text/html" in content_type and status_code == 200:
            soup = BeautifulSoup(r.text, "html.parser")
            title_tag = soup.find("title")
            title = title_tag.get_text(strip=True) if title_tag else ""
            
            md_tag = soup.find("meta", attrs={"name": "description"})
            meta_desc = md_tag.get("content", "").strip() if md_tag else ""
            
            can_tag = soup.find("link", attrs={"rel": "canonical"})
            canonical = can_tag.get("href", "").strip() if can_tag else ""
            
            rm_tag = soup.find("meta", attrs={"name": "robots"})
            robots_meta = rm_tag.get("content", "").strip() if rm_tag else ""
            
            vp_tag = soup.find("meta", attrs={"name": "viewport"})
            viewport_found = bool(vp_tag)
            
            ogt = soup.find("meta", attrs={"property": "og:title"})
            og_title = ogt.get("content", "").strip() if ogt else ""
            
            ogi = soup.find("meta", attrs={"property": "og:image"})
            og_image = ogi.get("content", "").strip() if ogi else ""
            
            h1_list = [h.get_text(strip=True) for h in soup.find_all("h1")]
            h2_list = [h.get_text(strip=True) for h in soup.find_all("h2")]
            
            imgs = soup.find_all("img")
            total_images = len(imgs)
            images_missing_alt = sum(1 for img in imgs if not img.get("alt"))
            
            scripts = soup.find_all("script", attrs={"type": "application/ld+json"})
            for s in scripts:
                try:
                    js = json.loads(s.string)
                    if isinstance(js, dict):
                        schema_types.append(js.get("@type", "Unknown"))
                    elif isinstance(js, list):
                        for item in js:
                            if isinstance(item, dict):
                                schema_types.append(item.get("@type", "Unknown"))
                except:
                    pass

        results.append({
            "url": url,
            "status_code": status_code,
            "latency_ms": latency_ms,
            "content_type": content_type,
            "size_bytes": content_len,
            "title": title,
            "title_length": len(title),
            "meta_description": meta_desc,
            "meta_description_length": len(meta_desc),
            "canonical": canonical,
            "robots_meta": robots_meta,
            "viewport_configured": viewport_found,
            "h1_count": len(h1_list),
            "h1_content": " | ".join(h1_list),
            "h2_count": len(h2_list),
            "total_images": total_images,
            "images_missing_alt": images_missing_alt,
            "og_title": og_title,
            "og_image": og_image,
            "schema_types": ", ".join(schema_types) if schema_types else "None",
            "checked_at": timestamp_now
        })
        print(f"[{status_code}] {url} ({latency_ms}ms)")
    except Exception as e:
        results.append({
            "url": url,
            "status_code": "ERROR",
            "error": str(e),
            "checked_at": timestamp_now
        })
        print(f"[ERR] {url}: {e}")

# Save json evidence
evidence_path = r"D:\incubrix project final\technical_audit\audit_evidence.json"
with open(evidence_path, "w", encoding="utf-8") as f:
    json.dump(results, f, indent=2)

# Save audit results CSV
csv_path = r"D:\incubrix project final\technical_audit\audit_results.csv"
fieldnames = list(results[0].keys())
with open(csv_path, "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(results)

print("Saved audit_results.csv and audit_evidence.json.")
