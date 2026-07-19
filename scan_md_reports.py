import os
import re

dir_path = r"C:\Users\user\Desktop\Antigravity\データ分析自動化\data_input"

for f in os.listdir(dir_path):
    if "ai_report.md" in f:
        path = os.path.join(dir_path, f)
        print(f"\n=============================")
        print(f"File: {f}")
        with open(path, "r", encoding="utf-8") as file:
            content = file.read()
            
        table_pattern = r'\|?\s*(?:\d{4}/\d{2}/\d{2})\s*\|.*'
        table_lines = re.findall(table_pattern, content)
        print(f"Found {len(table_lines)} table lines:")
        for line in table_lines:
            print(f"  {line}")
