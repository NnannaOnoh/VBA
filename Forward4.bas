Attribute VB_Name = "Forward4"
Sub ForwardFacturen4()
     
    Dim oExplorer As Outlook.Explorer
    Dim oMail As MailItem
    Set oExplorer = Application.ActiveExplorer
    Set oMail = oExplorer.Selection.item(1).Forward
    Dim i As Variant
    
    i = InputBox("Factuurnummer")
     
    On Error GoTo Release
     
    If oExplorer.Selection.item(1).Class = olMail Then
        oMail.Subject = oMail.Subject & " Factuurnummer: " & i
        'oMail.HTMLBody = "Custom Text.<p> <img src=""custom image link""" _
        & " title=""D"" alt=""D"" name=""D"" border=""0"" id=""D""/>" _
        & vbCrLf & oMail.HTMLBody
        oMail.SentOnBehalfOfName = "facturen@amsterdam.nl"
        oMail.Recipients.Add "srvc47ACAM@amsterdam.nl"
        oMail.Recipients.item(1).Resolve
        If oMail.Recipients.item(1).Resolved Then
            oMail.Display
            'oMail.Save
            'oMail.Send
        Else
            MsgBox "Could not resolve " & oMail.Recipients.item(1).Address
        End If
    Else
        MsgBox ("Not a mail item")
    End If
Release:
    Set oMail = Nothing
    Set oExplorer = Nothing
    
    Call Afgehandeld
    
End Sub






