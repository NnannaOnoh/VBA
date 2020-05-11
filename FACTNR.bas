Attribute VB_Name = "FACTNR"
Option Explicit
Public Sub FactuurNummerInPDFNaam()
    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As attachment
    Dim Report As String
     
    For Each selItem In Application.ActiveExplorer.Selection
        If selItem.Class = olMail Then
            Set aMail = selItem
            For Each aAttach In aMail.Attachments
                Report = Report & GetAttachmentInfo(aAttach)
                Report = Report & "; "
            Next
            Call FactuurNummerInPDFNaam2("", Report)
        End If
    Next
End Sub

Sub FactuurNummerInPDFNaam2(Title As String, Report As String)
  
    Dim fwd As Outlook.MailItem
    Dim itm As Object
    Dim strUser As String
    
strUser = Left(Environ("USERNAME"), 3)
      
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set fwd = itm.Forward
        
        Do Until fwd.Attachments.Count = 0
            fwd.Attachments.Remove (1)
        Loop
        
        fwd.SentOnBehalfOfName = "facturen@amsterdam.nl"
        fwd.Recipients.Add "srvc47ACAM@amsterdam.nl"
        
        CopyAttachments itm, fwd
        fwd.Subject = fwd.Subject & " " & Report
        fwd.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
        
        fwd.HTMLBody = fwd.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser
        
        fwd.Display
        fwd.DeferredDeliveryTime = DateAdd("s", 25, Now)
        
    End If
     
    Set fwd = Nothing
    Set itm = Nothing
    
    'Call Afgehandeld
    
End Sub
Public Sub FactuurNummerToevoegen()
    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As attachment
    Dim Report As String
     
    For Each selItem In Application.ActiveExplorer.Selection
        If selItem.Class = olMail Then
            Set aMail = selItem
            For Each aAttach In aMail.Attachments
                Report = Report & GetAttachmentInfo(aAttach)
                Report = Report & "; "
            Next
            Call FactuurNummerToevoegen2("", Report)
        End If
    Next
End Sub
Sub FactuurNummerToevoegen2(Title As String, Report As String)
  
    Dim fwd As Outlook.MailItem
    Dim itm As Object
    Dim i As String
    Dim strUser As String
    Dim pthBREAK As String
    Dim strBREAK As String
    
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

    strUser = Left(Environ("USERNAME"), 3)
    
    pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
    strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll
    
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
        fwd.HTMLBody = fwd.HTMLBody & Report & " " & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " " & strUser & strBREAK
        fwd.DeferredDeliveryTime = DateAdd("s", 25, Now)
        fwd.Send
        
    End If
     
    Set fwd = Nothing
    Set itm = Nothing
    
    Call Afgehandeld
    
End Sub
