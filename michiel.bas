Attribute VB_Name = "michiel"
Sub ForwardMichiel()
     
    Dim oExplorer As Outlook.Explorer
    Dim oMail As MailItem
    Set oExplorer = Application.ActiveExplorer
    Set oMail = oExplorer.Selection.item(1).Forward
    Dim i As Variant
    
    i = InputBox("Documentnummer")
     
    On Error GoTo Release
     
    If oExplorer.Selection.item(1).Class = olMail Then
        'oMail.Subject = "FW: Personalized Subject Line"
        oMail.HTMLBody = "document nummer: " & i '_
        '& " title=""D"" alt=""D"" name=""D"" border=""0"" id=""D""/>" _
        '& vbCrLf & oMail.HTMLBody
        oMail.SentOnBehalfOfName = "facturen@amsterdam.nl"
        oMail.Recipients.Add "M.van.der.Meulen@amsterdam.nl"
        oMail.Recipients.item(1).Resolve
        If oMail.Recipients.item(1).Resolved Then
            'oMail.Display
            'oMail.Save
            oMail.Send
        Else
            MsgBox "Could not resolve " & oMail.Recipients.item(1).Address
        End If
    Else
        MsgBox ("Not a mail item")
    End If
Release:
    Set oMail = Nothing
    Set oExplorer = Nothing
    
    Call michiel2
    
End Sub

