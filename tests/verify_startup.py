import urllib.request
import json
import csv
import openpyxl

print("=== 1. TESTING STREAMLIT (PORT 8501) ===")
for url in ["http://127.0.0.1:8501/", "http://localhost:8501/", "http://127.0.0.1:8501/_stcore/health"]:
    try:
        with urllib.request.urlopen(url, timeout=5) as resp:
            print(f"{url} -> Status: {resp.status}")
    except Exception as e:
        print(f"{url} -> Error: {e}")

print("\n=== 2. TESTING FASTAPI (PORT 8000) ===")
for url in ["http://127.0.0.1:8000/", "http://127.0.0.1:8000/docs", "http://127.0.0.1:8000/api/health", "http://127.0.0.1:8000/api/questions?limit=2"]:
    try:
        with urllib.request.urlopen(url, timeout=5) as resp:
            data = resp.read()
            print(f"{url} -> Status: {resp.status}, Bytes: {len(data)}")
    except Exception as e:
        print(f"{url} -> Error: {e}")

print("\n=== 3. DATASET PRESERVATION & WORKBOOK INTEGRITY ===")
q_csv = r"D:\incubrix project final\data\creator_questions_100.csv"
with open(q_csv, "r", encoding="utf-8") as f:
    q_rows = list(csv.DictReader(f))
print(f"Questions CSV count: {len(q_rows)}")

wb_path = r"D:\incubrix project final\submission\IncuBrix_Option3_Submission_Workbook_COMPLETED.xlsx"
wb = openpyxl.load_workbook(wb_path, data_only=False)
ws_sc = wb["04_SELF_CHECK"]
print(f"Workbook 04_SELF_CHECK D6 formula: {ws_sc['D6'].value}")
print(f"Workbook 04_SELF_CHECK C29 formula: {ws_sc['C29'].value}")
print("Verification complete.")
