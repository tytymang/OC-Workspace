import win32com.client
import json

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)

items = inbox.Items
items.Sort("[ReceivedTime]", True)

data = []
count = 0
for item in items:
    if count >= 15: break
    if item.UnRead and "Dataiku" in item.Subject:
        data.append({
            "Sender": item.SenderName,
            "Subject": item.Subject,
            "HTMLBody": item.HTMLBody,
            "Time": item.ReceivedTime.strftime("%m-%d %H:%M")
        })
    count += 1

with open(r'C:\Users\307984\.openclaw\workspace\emails_python2.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=4)
