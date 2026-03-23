import win32com.client as win32
import json

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
wb = excel.Workbooks.Open(r"C:\Users\307984\.openclaw\document\IT_AI_과제_취합.xlsx")
ws = wb.Worksheets("Sheet1")

data = []
for r in range(4, 10):
    row_data = []
    for c in range(1, 30):
        val = ws.Cells(r, c).Value
        row_data.append(str(val).replace('\n', ' ') if val else "")
    data.append(row_data)

wb.Close(False)
excel.Quit()

with open(r"C:\Users\307984\.openclaw\workspace\source_sample.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
