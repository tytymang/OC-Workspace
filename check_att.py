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
        attachments = []
        for att in item.Attachments:
            attachments.append(att.FileName)
        data.append({
            "Sender": item.SenderName,
            "Subject": item.Subject,
            "Attachments": attachments
        })
    count += 1

print(json.dumps(data, ensure_ascii=False, indent=4))
