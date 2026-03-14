import win32com.client
import datetime
import json
import sys

def find_unreplied():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)
    sent_folder = ns.GetDefaultFolder(5)

    last_week = datetime.datetime.now() - datetime.timedelta(days=7)
    
    # Get Sent Items subjects
    sent_items = sent_folder.Items
    sent_items.Sort("[SentOn]", True)
    sent_subjects = set()
    for s in sent_items:
        try:
            if s.SentOn.replace(tzinfo=None) < last_week:
                break
            sub = s.Subject.replace("RE: ", "").replace("FW: ", "").strip()
            sent_subjects.add(sub)
        except:
            continue

    # Check Inbox
    received_items = inbox.Items
    received_items.Sort("[ReceivedTime]", True)
    
    results = []
    keywords = ["?", "부탁", "회신", "확인", "요청", "검토", "컨펌", "의견"]
    
    for m in received_items:
        try:
            if m.ReceivedTime.replace(tzinfo=None) < last_week:
                break
            
            # Skip if user is sender (sent to self)
            if m.SenderName == "나여나" or "307984" in m.SenderEmailAddress:
                continue

            sub = m.Subject.replace("RE: ", "").replace("FW: ", "").strip()
            if sub not in sent_subjects:
                body = m.Body
                if any(k in body for k in keywords) or m.Importance == 2:
                    results.append({
                        "Received": m.ReceivedTime.strftime("%m-%d %H:%M"),
                        "Sender": m.SenderName,
                        "Subject": m.Subject,
                        "Importance": "High" if m.Importance == 2 else "Normal"
                    })
        except:
            continue
            
    return results

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding='utf-8')
    unreplied = find_unreplied()
    print(json.dumps(unreplied, ensure_ascii=False, indent=2))
