import json
import re

def extract(html_path):
    with open(html_path, 'r', encoding='utf-8', errors='replace') as f:
        html = f.read()
    
    tables = re.findall(r'<table[^>]*>(.*?)</table>', html, re.IGNORECASE | re.DOTALL)
    results = []
    if tables:
        for table in tables:
            rows = re.findall(r'<tr[^>]*>(.*?)</tr>', table, re.IGNORECASE | re.DOTALL)
            for row in rows:
                cols = re.findall(r'<td[^>]*>(.*?)</td>', row, re.IGNORECASE | re.DOTALL)
                row_data = []
                for col in cols:
                    text = re.sub(r'<[^>]+>', '', col)
                    text = text.replace('\r', '').replace('\n', '').strip()
                    text = re.sub(r'\s+', ' ', text)
                    row_data.append(text)
                results.append(row_data)
    return results

data_0 = extract(r'C:\Users\307984\.openclaw\workspace\email_0.html')
with open(r'C:\Users\307984\.openclaw\workspace\table_0.json', 'w', encoding='utf-8') as f:
    json.dump(data_0, f, ensure_ascii=False, indent=4)
