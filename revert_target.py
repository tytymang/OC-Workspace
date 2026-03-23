import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
wb = excel.Workbooks.Open(r"C:\Users\307984\.openclaw\document\최종_전사_AI_과제_마스터_리스트.xlsx")
ws = wb.Worksheets("Sheet1")
lr = ws.UsedRange.Rows.Count
if lr > 22:
    ws.Range(ws.Cells(23, 1), ws.Cells(lr, ws.UsedRange.Columns.Count)).EntireRow.Delete()
    wb.Save()
    print(f"Reverted to 22 rows. Deleted {lr - 22} rows.")
else:
    print(f"No need to revert. Rows: {lr}")

wb.Close(False)
excel.Quit()
