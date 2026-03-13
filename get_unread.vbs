Dim oApp, oNS, oFolder, oItems, oItem, i, resultText
Set oApp = CreateObject("Outlook.Application")
Set oNS = oApp.GetNamespace("MAPI")
Set oFolder = oNS.GetDefaultFolder(6) ' 6 = Inbox
Set oItems = oFolder.Items
oItems.Sort "[ReceivedTime]", True
oItems.Restrict "[UnRead] = True"

resultText = "Subject|Sender|ReceivedTime" & vbCrLf
i = 0
For Each oItem In oItems
    If oItem.UnRead = True Then
        resultText = resultText & oItem.Subject & "|" & oItem.SenderName & "|" & oItem.ReceivedTime & vbCrLf
        i = i + 1
        If i >= 10 Then Exit For
    End If
Next

Dim fso, f
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile("unread_emails.txt", True, True) ' True for Unicode
f.WriteLine resultText
f.Close
