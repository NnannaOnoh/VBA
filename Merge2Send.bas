Attribute VB_Name = "Merge2Send"
Sub MergeToSend()
     
    Dim oExplorer As Outlook.Explorer
    Dim omail As MailItem
    Set oExplorer = Application.ActiveExplorer
    Dim itm As Object
    
    SaveAtt
    
    SaveToMerge
End Sub

Sub ForwardMerge3()

    Dim oExplorer As Outlook.Explorer
    Dim omail As MailItem
    Set oExplorer = Application.ActiveExplorer
    Dim itm As Object

    Set omail = oExplorer.Selection.Item(1).Forward
    
     Path = "H:\Mijn Documenten\merge\pdf\OLAttachments\"
     Ext = ".pdf"
    
    FN = InputBox("Factuurnummer")
    Select Case StrPtr(FN)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
    On Error GoTo Release
     
    If oExplorer.Selection.Item(1).Class = olMail Then
        
        'oMail.HTMLBody = "Custom Text.<p> <img src=""custom image link""" _
        & " title=""D"" alt=""D"" name=""D"" border=""0"" id=""D""/>" _
        & vbCrLf & oMail.HTMLBody
        omail.SentOnBehalfOfName = "facturen@amsterdam.nl"
        omail.Recipients.Add "srvc47ACAM@amsterdam.nl"
        omail.Recipients.Item(1).Resolve
        
        Do Until omail.Attachments.Count = 0
            omail.Attachments.Remove (1)
        Loop

        omail.Attachments.Add (Path & FN & "(M)" & ".pdf")
        omail.Subject = omail.Subject & " " & FN
        'If omail.Recipients.item(1).Resolved Then
            'omail.Display
            'oMail.Save
            omail.Send
        'Else
        '    MsgBox "Could not resolve " & omail.Recipients.item(1).Address
        'End If
    Else
        MsgBox ("Not a mail item")
    End If
Release:
    Set omail = Nothing
    Set oExplorer = Nothing
    
'    Application.Wait Second(Now) + 15
    
    'Kill (Path & "*.jpg")
    Kill (Path & "*.*")
    
    

    
    Call Afgehandeld
    
End Sub


Sub KillAll()

Kill ("H:\Mijn Documenten\merge\pdf\OLAttachments\*.*")

End Sub

Sub SplitPerPDF()

    Dim oExplorer As Outlook.Explorer
    Dim omail As MailItem
    Set oExplorer = Application.ActiveExplorer
    Dim itm As Object
    
    SaveAtt

    Set omail = oExplorer.Selection.Item(1).Forward
    
     Path = "H:\Mijn Documenten\merge\pdf\OLAttachments\"
     Ext = ".pdf"
     
BN = InputBox("Bestandsnaam")
Select Case StrPtr(BN)
    Case 0
    MsgBox ("Geannuleerd")
        Exit Sub
    Case Else
End Select
    
    FN = InputBox("Factuurnummer")
    Select Case StrPtr(FN)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
    On Error GoTo Release
     
    If oExplorer.Selection.Item(1).Class = olMail Then
        
        'oMail.HTMLBody = "Custom Text.<p> <img src=""custom image link""" _
        & " title=""D"" alt=""D"" name=""D"" border=""0"" id=""D""/>" _
        & vbCrLf & oMail.HTMLBody
        omail.SentOnBehalfOfName = "facturen@amsterdam.nl"
        omail.Recipients.Add "srvc47ACAM@amsterdam.nl"
        omail.Recipients.Item(1).Resolve
        
        Do Until omail.Attachments.Count = 0
            omail.Attachments.Remove (1)
        Loop

        omail.Attachments.Add (Path & BN & ".pdf")
        omail.Subject = omail.Subject & " " & FN
        'If omail.Recipients.item(1).Resolved Then
            'omail.Display
            'oMail.Save
            omail.Send
        'Else
        '    MsgBox "Could not resolve " & omail.Recipients.item(1).Address
        'End If
    Else
        MsgBox ("Not a mail item")
    End If
Release:
    Set omail = Nothing
    Set oExplorer = Nothing
    
'    Application.Wait Second(Now) + 15
    
    'Kill (Path & "*.jpg")
    Kill (Path & "*.*")
       
    'Call Afgehandeld
    
End Sub

