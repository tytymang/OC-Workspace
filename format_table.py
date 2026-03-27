import json
import sys

def mask_name(name):
    if not name: return ""
    if len(name) <= 2:
        return name[0] + "*"
    return name[0] + "*" * (len(name) - 2) + name[-1]

def mask_id(emp_id):
    if not emp_id or len(emp_id) < 4: return emp_id
    return emp_id[:2] + "**" + emp_id[4:]

with open(r'C:\Users\307984\.openclaw\workspace\attendee_data.json', 'r', encoding='utf-8-sig') as f:
    data = json.load(f)

with open(r'C:\Users\307984\.openclaw\workspace\attendee_table.md', 'w', encoding='utf-8') as out:
    out.write("| 순번 | 소속/팀 | 사번 | 이름 |\n")
    out.write("| :--- | :--- | :--- | :--- |\n")
    for i, item in enumerate(data, 1):
        team = item.get('Team', '') if item.get('Team') else "-"
        emp_id = mask_id(item.get('ID', ''))
        name = mask_name(item.get('Name', ''))
        out.write(f"| {i} | {team} | {emp_id} | {name} |\n")
