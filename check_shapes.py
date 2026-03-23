import win32com.client as win32
import json

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
wb = excel.Workbooks.Open(r"C:\Users\307984\.openclaw\document\IT_AI_과제_취합.xlsx")
ws = wb.Worksheets("Sheet1")

data = []
for r in range(4, 10):
    row_data = []
    # Check shape text in the row if possible?
    # Text boxes are shapes
    shapes_in_row = []
    for shp in ws.Shapes:
        try:
            if shp.TopLeftCell.Row == r or shp.BottomRightCell.Row == r:
                if shp.TextFrame2.HasText:
                    shapes_in_row.append(shp.TextFrame2.TextRange.Text.replace('\n', ' '))
        except Exception as e:
            pass
    data.append({"row": r, "shapes": shapes_in_row})

wb.Close(False)
excel.Quit()

with open(r"C:\Users\307984\.openclaw\workspace\shapes_sample.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
