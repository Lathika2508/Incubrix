import sys
import os

# Add root directory to sys.path
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from backend.app.database import init_db
from backend.app.engine import import_baseline_data, generate_weekly_report

def main():
    print("=" * 60)
    print("IncuBrix Visibility Engine - Dataset Initializer & Sync")
    print("=" * 60)
    init_db()
    import_baseline_data()
    print("Generating baseline weekly report...")
    rep = generate_weekly_report(1)
    print(f"Weekly baseline report generated: {rep}")
    print("Initialization completed successfully.")

if __name__ == "__main__":
    main()
