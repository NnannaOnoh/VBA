Attribute VB_Name = "FACTSCRIPTS"
Option Explicit
Public Sub FAC01_Onvolledige_Factuur()
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
            Call FAC01_Onvolledige_Factuur2("", Report)
        End If
    Next
End Sub

Sub FAC01_Onvolledige_Factuur2(Title As String, Report As String)

    Dim i As String
    Dim b As String
    Dim pthSig As String
    Dim strSig As String
    Dim pthBREAK As String
    Dim strBREAK As String
    Dim strUser As String
    
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)

pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 01.htm"
strSig = FSO.OpenTextFile(pthSig).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

    Dim rpl As Outlook.MailItem
    Dim itm As Object
    Dim verzonden As Integer
         
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set rpl = itm.Reply
        CopyAttachments itm, rpl
        
        i = InputBox("Factuurnummer")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
        b = InputBox("Bedrijfsnaam")
    Select Case StrPtr(b)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select

rpl.SentOnBehalfOfName = "facturen@amsterdam.nl"
rpl.Recipients.Item(1).Resolve
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
rpl.Attachments.Add ("G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\archief\flyer_1_administratie_a5_digitaal.pdf")
rpl.Subject = "Teruggestuurd/" & i & "/" & b & "/AE"
'rpl.Recipients.Add ""

rpl.HTMLBody = strSig & rpl.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK
                      
rpl.Display
rpl.DeferredDeliveryTime = DateAdd("s", 25, Now)

    End If
 
    Set rpl = Nothing
    Set itm = Nothing
    
    Call Retour

    End Sub
Public Sub FAC02_Campagne()
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
            Call FAC02_Campagne2("", Report)
        End If
    Next
End Sub

Sub FAC02_Campagne2(Title As String, Report As String)

    Dim i As String
    Dim b As String
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
    
    Dim rCount As String
    
    Dim strColA, strColB, strColC, strColD As String
       
    
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)

pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 02.htm"
strSig = FSO.OpenTextFile(pthSig).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

strPath = "G:\FIN\11DebCred\Crediteuren\60. Team Input\TIJDELIJK\Adresserings Campagne.xlsx"

    Dim rpl As Outlook.MailItem
    Dim itm As Object
         
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set rpl = itm.Reply
        CopyAttachments itm, rpl
        
        i = InputBox("Factuurnummer")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
        b = InputBox("Bedrijfsnaam")
    Select Case StrPtr(b)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select

rpl.SentOnBehalfOfName = "facturen@amsterdam.nl"
rpl.Recipients.Item(1).Resolve
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
rpl.Attachments.Add ("G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\archief\flyer_1_administratie_a5_digitaal.pdf")
'rpl.Recipients.Add ""

rpl.Subject = "Attentie/" & i & "/" & b & "/AE"

rpl.HTMLBody = strSig & rpl.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK
rpl.Display
rpl.DeferredDeliveryTime = DateAdd("s", 25, Now)

    End If
 
    Set rpl = Nothing
    Set itm = Nothing
    
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
    
    On Error Resume Next
rCount = xlSheet.Range("B" & xlSheet.Rows.Count).End(-4162).Row
rCount = rCount + 1

    strColA = strUser
    strColB = (Format(Now, "yyyy-mm-dd HH:MM:SS"))
    strColC = b
    strColD = i
    'strColE = ""
    'strColF = ""
    'strColG = ""
    'strColH = ""

  xlSheet.Range("a" & rCount) = strColA
  xlSheet.Range("B" & rCount) = strColB
  xlSheet.Range("c" & rCount) = strColC
  xlSheet.Range("d" & rCount) = strColD
  'xlSheet.Range("d" & rCount) = strColD
  'xlSheet.Range("e" & rCount) = strColE
  'xlSheet.Range("f" & rCount) = strColF
  'xlSheet.Range("g" & rCount) = strColG
  'xlSheet.Range("h" & rCount) = strColH

  rCount = rCount + 1
 
Application.DisplayAlerts = False
 
xlWb.Save '("G:\FIN\11DebCred\Crediteuren\60. Team Input\TIJDELIJK\Adresserings Campagne.xlsx")
xlWb.Close False
Application.DisplayAlerts = False


     'If bXStarted Then
     '    xlApp.Quit
     'End If

     Set xlApp = Nothing
     Set xlWb = Nothing
     Set xlSheet = Nothing

    'Call Retour
    
'Workbooks("VERWIJDERD UIT KOFAX1.xlsm").Close False
'Workbook.Close ("KOFAX.xlsx")

'Workbooks("VERWIJDERD UIT KOFAX.xlsm").Sheets(1).Range("A2") = ""
'Workbooks("VERWIJDERD UIT KOFAX.xlsm").Sheets(1).Range("B2") = ""
'Workbooks("VERWIJDERD UIT KOFAX.xlsm").Sheets(1).Range("C2") = ""
'Workbooks("VERWIJDERD UIT KOFAX.xlsm").Sheets(1).Range("D2") = ""
'Workbooks("VERWIJDERD UIT KOFAX.xlsm").Sheets(1).Range("E2") = ""
'Workbooks("VERWIJDERD UIT KOFAX.xlsm").Sheets(1).Range("F2") = ""
'Workbooks("VERWIJDERD UIT KOFAX.xlsm").Sheets(1).Range("G2") = ""

    End Sub
Public Sub FAC03_Onvolledige_Creditfactuur()
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
            Call FAC03_Onvolledige_Creditfactuur2("", Report)
        End If
    Next
End Sub
Sub FAC03_Onvolledige_Creditfactuur2(Title As String, Report As String)

    Dim i As String
    Dim b As String
    Dim pthSig As String
    Dim strSig As String
    Dim pthBREAK As String
    Dim strBREAK As String
    Dim strUser As String
    
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)

pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 03.htm"
strSig = FSO.OpenTextFile(pthSig).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

    Dim rpl As Outlook.MailItem
    Dim itm As Object
    Dim verzonden As Integer
         
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set rpl = itm.Reply
        CopyAttachments itm, rpl
        
        i = InputBox("Factuurnummer")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
        b = InputBox("Bedrijfsnaam")
    Select Case StrPtr(b)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select

rpl.SentOnBehalfOfName = "facturen@amsterdam.nl"
rpl.Recipients.Item(1).Resolve
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\archief\flyer_1_administratie_a5_digitaal.pdf"
'rpl.Recipients.Add ""
        
rpl.Subject = "Teruggestuurd/" & i & "/" & b & "/CR; AE"

rpl.HTMLBody = strSig & rpl.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK
rpl.Display
rpl.DeferredDeliveryTime = DateAdd("s", 25, Now)

    End If
 
    Set rpl = Nothing
    Set itm = Nothing
    
    Call Retour

    End Sub
Public Sub FAC04_Betaalopdracht_In_Facturen()
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
            Call FAC04_Betaalopdracht_In_Facturen2("", Report)
        End If
    Next
End Sub
Sub FAC04_Betaalopdracht_In_Facturen2(Title As String, Report As String)

    Dim i As String
    Dim b As String
    Dim pthSig As String
    Dim strSig As String
    Dim pthBREAK As String
    Dim strBREAK As String
    Dim strUser As String
    
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)

pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 04.htm"
strSig = FSO.OpenTextFile(pthSig).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

    Dim rpl As Outlook.MailItem
    Dim itm As Object
    Dim verzonden As Integer
         
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set rpl = itm.Reply
        CopyAttachments itm, rpl
        
        i = InputBox("Factuurnummer")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
        b = InputBox("Bedrijfsnaam")
    Select Case StrPtr(b)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select

rpl.SentOnBehalfOfName = "facturen@amsterdam.nl"
rpl.Recipients.Item(1).Resolve
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\archief\flyer_1_administratie_a5_digitaal.pdf"
'rpl.Recipients.Add ""
        
rpl.Subject = "Teruggestuurd/" & i & "/" & b & "/AE"

rpl.HTMLBody = strSig & rpl.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK
rpl.Display
rpl.DeferredDeliveryTime = DateAdd("s", 25, Now)

    End If
 
    Set rpl = Nothing
    Set itm = Nothing
    
    Call Retour

    End Sub
Public Sub FAC05_Factuur_Van_Collega()
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
            Call FAC05_Factuur_Van_Collega2("", Report)
        End If
    Next
End Sub
Sub FAC05_Factuur_Van_Collega2(Title As String, Report As String)

    Dim i As String
    Dim b As String
    Dim pthSig As String
    Dim strSig As String
    Dim pthBREAK As String
    Dim strBREAK As String
    Dim strUser As String
    
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)

pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 05.htm"
strSig = FSO.OpenTextFile(pthSig).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

    Dim rpl As Outlook.MailItem
    Dim itm As Object
    Dim verzonden As Integer
         
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set rpl = itm.Reply
        CopyAttachments itm, rpl
        
        i = InputBox("Factuurnummer")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
        b = InputBox("Bedrijfsnaam")
    Select Case StrPtr(b)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select

rpl.SentOnBehalfOfName = "facturen@amsterdam.nl"
rpl.Recipients.Item(1).Resolve
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\archief\flyer_1_administratie_a5_digitaal.pdf"
'rpl.Recipients.Add ""
        
rpl.Subject = "Teruggestuurd/" & i & "/" & b & "/AE"

rpl.HTMLBody = strSig & rpl.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK
rpl.Display
rpl.DeferredDeliveryTime = DateAdd("s", 25, Now)

    End If
 
    Set rpl = Nothing
    Set itm = Nothing
    
    Call Retour

    End Sub
Public Sub FAC06_Vraag_Van_Collega()

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
            Call FAC06_Vraag_Van_Collega2("", Report)
        End If
    Next
End Sub
Sub FAC06_Vraag_Van_Collega2(Title As String, Report As String)

    Dim i As String
    Dim b As String
    Dim pthSig As String
    Dim strSig As String
    Dim pthBREAK As String
    Dim strBREAK As String
    Dim strUser As String
    
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)

pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 06.htm"
strSig = FSO.OpenTextFile(pthSig).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

    Dim rpl As Outlook.MailItem
    Dim itm As Object
    Dim verzonden As Integer
         
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set rpl = itm.Reply
        CopyAttachments itm, rpl
        
rpl.SentOnBehalfOfName = "facturen@amsterdam.nl"
rpl.Recipients.Item(1).Resolve
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\archief\flyer_1_administratie_a5_digitaal.pdf"
'rpl.Recipients.Add ""
        
rpl.Subject = "Uw vraag wordt niet in behandeling genomen"

rpl.HTMLBody = strSig & rpl.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK
rpl.Display
rpl.DeferredDeliveryTime = DateAdd("s", 25, Now)

    End If
 
    Set rpl = Nothing
    Set itm = Nothing
    
    Call Retour

    End Sub
Public Sub FAC07_Incomplete_Invoice()
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
            Call FAC07_Incomplete_Invoice2("", Report)
        End If
    Next
End Sub
Sub FAC07_Incomplete_Invoice2(Title As String, Report As String)

    Dim i As String
    Dim b As String
    Dim pthSig As String
    Dim strSig As String
    Dim pthBREAK As String
    Dim strBREAK As String
    Dim strUser As String
    
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)

pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 07.htm"
strSig = FSO.OpenTextFile(pthSig).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

    Dim rpl As Outlook.MailItem
    Dim itm As Object
    Dim verzonden As Integer
         
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set rpl = itm.Reply
        CopyAttachments itm, rpl
        
        i = InputBox("Factuurnummer")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
        b = InputBox("Bedrijfsnaam")
    Select Case StrPtr(b)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select

rpl.SentOnBehalfOfName = "facturen@amsterdam.nl"
rpl.Recipients.Item(1).Resolve
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\archief\flyer_1_administratie_a5_digitaal.pdf"
'rpl.Recipients.Add ""
        
rpl.Subject = "Teruggestuurd/" & i & "/" & b & "/AE"

rpl.HTMLBody = strSig & rpl.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK
rpl.Display
rpl.DeferredDeliveryTime = DateAdd("s", 25, Now)

    End If
 
    Set rpl = Nothing
    Set itm = Nothing
    
    Call Retour

    End Sub
Public Sub FAC08_Vraag_Van_Leverancier()
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
            Call FAC08_Vraag_Van_Leverancier2("", Report)
        End If
    Next
End Sub
Sub FAC08_Vraag_Van_Leverancier2(Title As String, Report As String)

    Dim i As String
    Dim b As String
    Dim pthSig As String
    Dim strSig As String
    Dim pthBREAK As String
    Dim strBREAK As String
    Dim strUser As String
    
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)

pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 08.htm"
strSig = FSO.OpenTextFile(pthSig).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

    Dim rpl As Outlook.MailItem
    Dim itm As Object
    Dim verzonden As Integer
         
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set rpl = itm.Reply
        CopyAttachments itm, rpl
        
rpl.SentOnBehalfOfName = "facturen@amsterdam.nl"
rpl.Recipients.Item(1).Resolve
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\archief\flyer_1_administratie_a5_digitaal.pdf"
'rpl.Recipients.Add ""
        
rpl.Subject = "Uw vraag wordt niet in behandeling genomen"

rpl.HTMLBody = strSig & rpl.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK
rpl.Display
rpl.DeferredDeliveryTime = DateAdd("s", 25, Now)

    End If
 
    Set rpl = Nothing
    Set itm = Nothing
    
    Call Retour

    End Sub
Sub credit()

RedenRetour.RC = False
RedenRetour.PDF = False
RedenRetour.IBAN = False

RedenRetour.DEB = False
RedenRetour.NAWLEV = False
RedenRetour.NAWGA = False
RedenRetour.FCTNR = False
RedenRetour.FCTDTM = False
RedenRetour.KVK = False
RedenRetour.BTW = False
RedenRetour.BDRG = False
RedenRetour.DNB = False

Kill ("H:\Mijn Documenten\merge\pdf\OLAttachments\watermerk\*.*")

RedenRetour.Show

End Sub
