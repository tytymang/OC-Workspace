import win32com.client
import sys
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
items = inbox.Items
items.Sort("[ReceivedTime]", True)
for item in items:
    if "Dataiku" in item.Subject and "JaeBeom" in item.SenderName:
        with open(r'C:\Users\307984\.openclaw\workspace\vp_email_text.txt', 'w', encoding='utf-8') as f:
            f.write(item.Body)
        break
