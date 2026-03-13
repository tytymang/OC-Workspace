Set objOutlook = CreateObject("Outlook.Application")
If TypeName(objOutlook) = "Application" Then
    WScript.Echo "Outlook Object Created"
    On Error Resume Next
    Set objVBE = objOutlook.VBE
    If Err.Number <> 0 Then
        WScript.Echo "VBE Error: " & Err.Description
    ElseIf objVBE Is Nothing Then
        WScript.Echo "VBE is Nothing (Macro Security Blocks Access)"
    Else
        WScript.Echo "VBE is accessible."
    End If
End If
