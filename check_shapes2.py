import win32com.client as win32
import json

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
wb = excel.Workbooks.Open(r"C:\Users\307984\.openclaw\document\IT_AI_과제_취합.xlsx")
ws = wb.Worksheets("Sheet1")

data = []
for r in range(4, 10):
    row_data = []
    shapes_in_row = []
    for shp in ws.Shapes:
        try:
            # We want shapes that vertically intersect this row
            if shp.Top >= ws.Rows(r).Top and shp.Top < ws.Rows(r+1).Top:
                if shp.TextFrame2.HasText:
                    text = shp.TextFrame2.TextRange.Text.replace('\n', ' ')
                    left = shp.Left
                    shapes_in_row.append({"text": text.strip(), "left": left})
        except Exception as e:
            pass
    
    # sort shapes by left position
    shapes_in_row.sort(key=lambda x: x["left"])
    sorted_texts = [s["text"] for s in shapes_in_row]
    data.append({"row": r, "shapes": sorted_texts})

wb.Close(False)
excel.Quit()

with open(r"C:\Users\307984\.openclaw\workspace\shapes_sample2.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
