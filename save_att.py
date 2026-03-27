import win32com.client
import os

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
items = inbox.Items
items.Sort("[ReceivedTime]", True)

for item in items:
    if "Dataiku" in item.Subject and item.UnRead:
        if "MyeongGyun" in item.SenderName or "Choi" in item.SenderName:
            for att in item.Attachments:
                path = os.path.join(r"C:\Users\307984\.openclaw\workspace", att.FileName)
                att.SaveAsFile(path)
        if "Jaden" in item.SenderName:
            for att in item.Attachments:
                path = os.path.join(r"C:\Users\307984\.openclaw\workspace", "jaden_" + att.FileName)
                att.SaveAsFile(path)
