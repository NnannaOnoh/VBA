Attribute VB_Name = "FNToevoegen"
Sub FactuurNummerToevoegen()
  
    Dim fwd As Outlook.MailItem
    Dim itm As Object
    
i = InputBox("Factuurnummer")
Select Case StrPtr(i)
    Case 0
    MsgBox ("Geannuleerd")
        Exit Sub
    Case Else
End Select
  
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set fwd = itm.Forward
        
        Do Until fwd.Attachments.Count = 0
            fwd.Attachments.Remove (1)
        Loop
        
        fwd.SentOnBehalfOfName = "facturen@amsterdam.nl"
        fwd.Recipients.Add "srvc47ACAM@amsterdam.nl"
        
        CopyAttachments itm, fwd
        fwd.Subject = fwd.Subject & " Factuurnummer: " & i
        fwd.Send
    End If
     
    Set fwd = Nothing
    Set itm = Nothing
    
    Call Afgehandeld
    
End Sub
