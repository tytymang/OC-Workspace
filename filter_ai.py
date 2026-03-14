import re

with open(r'C:\Users\307984\.openclaw\workspace\extracted_data.txt', 'r', encoding='utf-8') as f:
    lines = f.readlines()

output = []
recording = False
for i, line in enumerate(lines):
    if "Sheet:" in line or "Slide" in line or "Opening PPT" in line:
        output.append("\n" + line.strip())
        recording = False
    
    # Heuristic for AI tasks in Excel
    if "AI과제" in line or "AI/시그나비오" in line or "AI활용" in line or "AI:" in line:
        recording = True
    
    if recording or "AI" in line:
        output.append(line.strip())

with open(r'C:\Users\307984\.openclaw\workspace\ai_tasks_filtered.txt', 'w', encoding='utf-8') as f:
    f.write("\n".join(output))

print("Filtered saved.")
