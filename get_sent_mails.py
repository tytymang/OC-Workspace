import win32com.client
import json
import sys

def get_recent_sent_mails(count=5):
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    sent_folder = ns.GetDefaultFolder(5) # olFolderSentMail
    items = sent_folder.Items
    items.Sort("[SentOn]", True)
    
    results = []
    for i in range(1, min(count + 1, items.Count + 1)):
        try:
            m = items.Item(i)
            results.append({
                "SentOn": m.SentOn.strftime("%Y-%m-%d %H:%M"),
                "To": m.To,
                "Subject": m.Subject
            })
        except:
            continue
    return results

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding='utf-8')
    sent_mails = get_recent_sent_mails()
    print(json.dumps(sent_mails, ensure_ascii=False, indent=2))
