import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
items = inbox.Items
items.Sort("[ReceivedTime]", True)
for item in items:
    if item.UnRead and "Dataiku" in item.Subject and "MyeongGyun" in item.SenderName:
        print(item.Body)
        break
