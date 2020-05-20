Attribute VB_Name = "HERRI"
'Public Sub Herinnering33()
'    Dim selItem As Object
'    Dim aMail As MailItem
'    Dim aAttach As attachment
'    Dim Report As String
'    Dim t As Date
'
'
'    For Each selItem In Application.ActiveExplorer.Selection
'        If selItem.Class = olMail Then
'            Set aMail = selItem
'            For Each aAttach In aMail.Attachments
'                Report = Report & GetAttachmentInfo(aAttach)
'                Report = Report & ", "
'            Next
'
'           ond = selItem.Subject
'             t = selItem.ReceivedTime
'            EM = selItem.SenderEmailAddress
'
'            Call Herinnering2("", Report, t, EM, ond, aMail)
'
'        End If
'    Next
'End Sub

'Sub Herinnering22(Title As String, Report As String, t, EM, ond, aMail)
'    Dim fwd As Outlook.MailItem
'    Dim itm As Object
'    Dim strUser As String
'
'    strUser = Left(Environ("USERNAME"), 3)
'
'    Set itm = GetCurrentItem()
'    If Not itm Is Nothing Then
'        Set fwd = itm.Forward
'
'        Do Until fwd.Attachments.Count = 0
'                 fwd.Attachments.Remove (1)
'        Loop
'
'        fwd.SentOnBehalfOfName = "facturen@amsterdam.nl"
'        fwd.Recipients.Add "crediteurenadministratie@amsterdam.nl"
'        fwd.Attachments.Add aMail, olEmbeddeditem
'        fwd.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
'        CopyAttachmentsFM itm, fwd
'        fwd.Subject = "Herinnering | Aanmaning ontvangen van " & EM
'        fwd.HTMLBody = "<p style=font-size:14px;font-family:corbel;color:black><br>Mail onderwerp: " & ond & "<br>" _
'                      & "E-mail ontvangen van: " & EM & " op " & t & "<br>" _
'                      & fwd.HTMLBody & "<br><p style=font-size:14px;font-family:corbel;color:black><br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
'                      & "<img src='cid:VHB.png' width='27' height='17'>" & "<br><p style=font-size:14px;font-family:corbel;color:white>" _
'                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser
'
'        fwd.DeferredDeliveryTime = DateAdd("s", 25, Now)
'        fwd.Send
'
'    End If
'
'    Set fwd = Nothing
'    Set itm = Nothing
'
'    Call KNOP7
'
'    Call Herrin
'
'End Sub

'Sub Herinnering2(Title As String, Report As String, t, EM, ond, aMail)
'    Dim fwd As Outlook.MailItem
'    Dim itm As Object
'    Dim strUser As String
'
'    strUser = Left(Environ("USERNAME"), 3)
'
'    Set itm = GetCurrentItem()
'    If Not itm Is Nothing Then
'        Set fwd = itm.Forward
'
'        Do Until fwd.Attachments.Count = 0
'            fwd.Attachments.Remove (1)
'        Loop
'
'        fwd.SentOnBehalfOfName = "facturen@amsterdam.nl"
'        fwd.Recipients.Add "crediteurenadministratie@amsterdam.nl"
'
'        CopyAttachments itm, fwd
'        fwd.Subject = fwd.Subject
'        fwd.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
'        fwd.HTMLBody = fwd.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
'                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
'                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser
'        fwd.DeferredDeliveryTime = DateAdd("s", 25, Now)
'        fwd.Send
'
'
'    End If
'
'    Set fwd = Nothing
'    Set itm = Nothing
'
'
'    Call KNOP7
'
'    Call Herrin
'
'End Sub

Sub Herinnering()
 
    Dim olNameSpace As Outlook.NameSpace
    Dim objCopyFolder As Outlook.MAPIFolder
    Dim objItem As Outlook.MailItem
    Dim objCopy As Outlook.MailItem
    
    Set olNameSpace = Application.GetNamespace("MAPI")

    Set objCopyFolder = olNameSpace.Folders("Facturen").Folders("Herinneringen & Aanmaningen")

    Set objItem = Application.ActiveExplorer.Selection.Item(1)

     Set objCopy = objItem.Copy
      objCopy.Move objCopyFolder

        With objItem
            .UnRead = True
            .Categories = "Naar CA"
            .Save
        End With
        
Afgehandeld

KNOP7
       
KillAll

    Set olNameSpace = Nothing
    Set objCopyFolder = Nothing
    Set objNS = Nothing
    Set objItem = Nothing
    Set objCopy = Nothing
 
End Sub

