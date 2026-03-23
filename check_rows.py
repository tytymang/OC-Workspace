import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
wb = excel.Workbooks.Open(r"C:\Users\307984\.openclaw\document\최종_전사_AI_과제_마스터_리스트.xlsx")
ws = wb.Worksheets("Sheet1")
print(f"Target file has {ws.UsedRange.Rows.Count} rows.")
wb.Close(False)
excel.Quit()
