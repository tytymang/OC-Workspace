import win32com.client
import os

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)

items = inbox.Items
items.Sort("[ReceivedTime]", True)

count = 0
for item in items:
    if count >= 15: break
    if item.UnRead and "Dataiku" in item.Subject:
        path = os.path.join(r"C:\Users\307984\.openclaw\workspace", f"email_{count}.html")
        item.SaveAs(path, 5) # 5 = olHTML
    count += 1
