Attribute VB_Name = "CRDTCOMP"
Option Explicit
Public Sub CreditCompleet()
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
            Call CreditCompleet2("", Report)
        End If
    Next
End Sub
Sub CreditCompleet2(Title As String, Report As String, FACTUURNUMMER As String)
  
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
        fwd.Recipients.Add "srvc18vr@amsterdam.nl"
        
        CopyAttachments itm, fwd
        fwd.Subject = fwd.Subject & FACTUURNUMMER
        fwd.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
        fwd.HTMLBody = fwd.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser
        
        fwd.DeferredDeliveryTime = DateAdd("s", 25, Now)
        fwd.Send
        
        
    End If
     
    Set fwd = Nothing
    Set itm = Nothing
    
    Call Afgehandeld
    
End Sub
Sub SRCHFL()

    Dim FACTUURNUMMER As String
    Dim b As String
    Dim Fl As String
    Dim Fd As String
    Dim Fb As String
    Dim pthSig As String
    Dim strSig As String
    Dim pthBREAK As String
    Dim strBREAK As String
    Dim strUser As String
    
    Dim strPath As String
    
    Dim xlApp As Object
    Dim xlWb As Workbook
    Dim xlSheet As Worksheet
    Dim bXStarted As Integer

           FACTUURNUMMER = InputBox("Factuurnummer")
    Select Case StrPtr(FACTUURNUMMER)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select

strPath = "C:\Users\Nnanna\Desktop\test.xlsx"
     On Error Resume Next
     Set xlApp = GetObject(, "Excel.Application")
     If Err <> 0 Then
         Application.StatusBar = "Please wait while Excel source is opened ... "
         Set xlApp = CreateObject("Excel.Application")
         bXStarted = True
     End If
     On Error GoTo 0

     Set xlWb = xlApp.Workbooks.Open(strPath)
     Set xlSheet = xlWb.Sheets(1)

            Cells.Find(What:=FACTUURNUMMER).Activate

            Selection.Offset(0, 5).Select

    Selection.Activate
    
    ActiveCell.Value2 = ("Afgehandeld")
            
            
            xlWb.Save
            xlWb.Close False
            
On Error Resume Next

End Sub
