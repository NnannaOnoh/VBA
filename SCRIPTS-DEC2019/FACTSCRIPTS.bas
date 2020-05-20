Attribute VB_Name = "FACTSCRIPTS"
Option Explicit
Public Sub FAC01_Onvolledige_Factuur()
    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As Attachment
    Dim Report As String
     
    For Each selItem In Application.ActiveExplorer.Selection
        If selItem.Class = olMail Then
            Set aMail = selItem
            For Each aAttach In aMail.Attachments
                Report = Report & GetAttachmentInfo(aAttach) '< Functie in ATTAC die ervoor zorgt dat de naam van de bijlage kan worden meegenomen naar de body van de text
                Report = Report & "; "
            Next
            Call FAC01_Onvolledige_Factuur2("", Report) '< ga naar "mail opmaken" met naam van de PDF (Report)
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

strUser = Left(Environ("USERNAME"), 3) '< gebruikersnaam voor onderaan de mail (eerste drie letters ADW account

pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 01.htm" '< pthSig locatie van htm bestand op de gedeelde schijf die de text voor de body van de mail bevat

strSig = FSO.OpenTextFile(pthSig).ReadAll '< strSig opdracht om text in de mail te zetten (komt in de opmaak van de mail weer terug)

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm" '< pthBREAK (lege regel)

strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll '< strBREAK opdracht om een lege regel in de mail te zetten (komt in de opmaak van de mail weer terug)

    Dim rpl As Outlook.MailItem
    Dim itm As Object
    Dim verzonden As Integer
         
    Set itm = GetCurrentItem() '< Functie in ATTAC die ervoor zorgt dat de module aan de in Outlook geselecteerde mail wordt gerelateerd
        
    If Not itm Is Nothing Then '< begin van de IF statement
        Set rpl = itm.Reply '< verzend als antwoord (Reply)
        CopyAttachments itm, rpl '< Functie in ATTAC die ervoor zorgt dat de bijlagen verwerkt worden
        
        
        i = InputBox("Factuurnummer") '< Input voor onderwerp van de mail
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
        b = InputBox("Bedrijfsnaam") '< Input voor onderwerp van de mail
    Select Case StrPtr(b)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select

rpl.SentOnBehalfOfName = "facturen@amsterdam.nl" '< "VAN"-veld
rpl.Recipients.Item(1).Resolve '< "AAN"-veld, het betreft hier een reply het e-mailadres wordt dus uit de "CurrentItem" opgemaakt "resolved"
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0 '< Vastberaden, Barmhartig, Heldhaftig (afbeeldingen onder aan de mail) positie is "0" dus niet zichtbaar als "bijlage"
'rpl.Attachments.Add "" <'ruimte voor bijlage (hier stond eerst de locatie van de "Financiën Brochure - Factuureisen"

rpl.Subject = "Teruggestuurd/" & i & "/" & b & "/AE" '< Het onderwerp van de mail

'rpl.Recipients.Add "" '> ruimte voor het toevoegen van een geadresseerde

'HTML code om mail body op te maken
         rpl.HTMLBody = "<p style=font-size:14px;font-family:corbel;color:black>" & strSig & rpl.HTMLBody _
                      & "<br><br>_____________________________________________________&nbsp;" _
                      & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK

rpl.DeferredDeliveryTime = DateAdd("s", 25, Now) '< verzend vertraging, mail blijft 25 seconden in Postvak UIT voor verzending
rpl.Display '< Display mail, verzend mail niet "blind"


    End If '< Eind van de IF statement
 
    Set rpl = Nothing
    Set itm = Nothing
    
    Call KNOP2 '< aansturing voor registratie medewerkersformulier (EXCELKNOP)
    
    Call Retour '< aansturing voor verplaatsen mail naar map retour

    End Sub
Public Sub FAC02_Campagne()

'Campagneknop met extra registratie van verwerking in Excel bestand

    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As Attachment
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
    Dim xlWB As Workbook
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
'rpl.Attachments.Add ""
'rpl.Recipients.Add ""

rpl.Subject = "Attentie/" & i & "/" & b & "/AE"

rpl.HTMLBody = strSig & rpl.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK

rpl.DeferredDeliveryTime = DateAdd("s", 25, Now)
rpl.Display


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
     Set xlWB = xlApp.Workbooks.Open(strPath)
     Set xlSheet = xlWB.Sheets(1)
    
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
 
xlWB.Save '("G:\FIN\11DebCred\Crediteuren\60. Team Input\TIJDELIJK\Adresserings Campagne.xlsx")
xlWB.Close False
Application.DisplayAlerts = False


     'If bXStarted Then
     '    xlApp.Quit
     'End If

     Set xlApp = Nothing
     Set xlWB = Nothing
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
    Dim aAttach As Attachment
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
rpl.Attachments.Add ""
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
    
     Call KNOP2
    
    Call Retour

    End Sub
Public Sub FAC04_Betaalopdracht_In_Facturen()
    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As Attachment
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
rpl.Attachments.Add ""
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
    
     Call KNOP2
    
    Call Retour

    End Sub
Public Sub FAC05_Factuur_Van_Collega()
    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As Attachment
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
'rpl.Attachments.Add ""
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
    
     Call KNOP2
    
    Call Retour

    End Sub
Public Sub FAC06_Vraag_Van_Collega()

    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As Attachment
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
rpl.Attachments.Add ""
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
    
     Call KNOP2
    
    Call Retour

    End Sub
Public Sub FAC07_Incomplete_Invoice()
    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As Attachment
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
rpl.Attachments.Add ""
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
    
     Call KNOP2
    
    Call Retour

    End Sub
Public Sub FAC08_Vraag_Van_Leverancier()
    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As Attachment
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
rpl.Attachments.Add ""
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
    
     Call KNOP2
    
    Call Retour

    End Sub
Sub credit(FN)

RedenRetour.i = FN
RedenRetour.b = ""
RedenRetour.FD = ""
RedenRetour.Fl = ""

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
RedenRetour.CreditCompleet = False

'Kill ("H:\Mijn Documenten\merge\pdf\OLAttachments\watermerk\*.*")

RedenRetour.Show

End Sub
Sub Factuur_Retour1()

    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As Attachment
    Dim Report As String
    Dim EM As String
    Dim BdNm, BN As String

     
    For Each selItem In Application.ActiveExplorer.Selection
        If selItem.Class = olMail Then
            Set aMail = selItem
            For Each aAttach In aMail.Attachments
                Report = Report & GetAttachmentInfo(aAttach)
                Report = Report & " "
 
    Next
     EM = selItem.SenderEmailAddress
     BN = AAPJEPUNTJE(BdNm)
    End If
Next


On Error GoTo err1
FacturenRetour.FN = Left(Report, Len(Report) - 5)
err1:
FacturenRetour.CbFN = False

FacturenRetour.BN = BN

FacturenRetour.EM = EM
FacturenRetour.CbEM = False

FacturenRetour.OVRG = ""

FacturenRetour.CbAE = False
FacturenRetour.CbWE = False
FacturenRetour.CbOR = False
FacturenRetour.CbFM = False
FacturenRetour.CbCR = False
FacturenRetour.CbRC = False

FacturenRetour.CbAEAD = False
FacturenRetour.CbAEIO = False
FacturenRetour.CbAERC = False
FacturenRetour.CbWEAD = False
FacturenRetour.CbFMPD = False
FacturenRetour.CbWEBT = False
FacturenRetour.CbWEFD = False
FacturenRetour.CbWEFN = False
FacturenRetour.CbWEIB = False
FacturenRetour.CbWEKV = False

'FacturenRetour.ObDSPLEML = True
'FacturenRetour.ObSNDEML = True

FacturenRetour.MultiPage1.Value = 0

FacturenRetour.Show

End Sub
Sub Grootboek01(FN, RC, BdNm, EM, SN, SubTxT, CbAFGH, DSPLEML, t, OVRG)

GrootboekRoutecode.SubTxT = SubTxT

GrootboekRoutecode.FN = FN
GrootboekRoutecode.BdNm = BdNm

GrootboekRoutecode.RC = RC

GrootboekRoutecode.t = t
GrootboekRoutecode.SN = SN
GrootboekRoutecode.EM = EM

GrootboekRoutecode.OVRG = OVRG

GrootboekRoutecode.GEENRC = True
GrootboekRoutecode.DSPLEML = DSPLEML
GrootboekRoutecode.CbAFGH = CbAFGH

GrootboekRoutecode.Show

End Sub
Sub Factuur_Compleet2()

    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As Attachment
    Dim Report As String
    Dim Report1 As String
    Dim SN As String
    Dim EM As String
    Dim SubTxT As String
    Dim BdNm As String
    Dim DFA As String
    Dim t As String
    Dim REC As String
   
    For Each selItem In Application.ActiveExplorer.Selection
        If selItem.Class = olMail Then
            Set aMail = selItem

            For Each aAttach In aMail.Attachments
            
            If UCase(Right(aAttach.FileName, 3)) = "PDF" Then
                
                Report = Report & Left(GetAttachmentInfo(aAttach), Len(GetAttachmentInfo(aAttach)) - 4) & " & "
                Report1 = Report1 & GetAttachmentInfo(aAttach) & " & "
                'Report1 = Report1 & Report & Left(Report, Len(Report) - 4)
                'GetAttachmentInfo (aAttach) & ".pdf / "

          End If
          
     BdNm = AAPJEPUNTJE(BdNm)
          
     Next
  
     SubTxT = selItem.Subject
     EM = selItem.SenderEmailAddress
     SN = selItem.SenderName
     t = selItem.ReceivedTime
     
    End If
Next

FACTCMPL1.SubTxT = SubTxT

FACTCMPL1.CbHA = False
FACTCMPL1.CbGB = False
FACTCMPL1.CbCRED = False


On Error GoTo err1

FACTCMPL1.FN = Left(Report, Len(Report) - 3)
FACTCMPL1.BN = Left(Report1, Len(Report1) - 3)

err1:

FACTCMPL1.CbMrg = False

FACTCMPL1.CbARC = False
FACTCMPL1.CbAEIO = False

FACTCMPL1.RC = ""
FACTCMPL1.t = t
FACTCMPL1.SN = SN

'FACTCMPL1.CbCpg = False
FACTCMPL1.BdNm = BdNm
FACTCMPL1.EM = EM

FACTCMPL1.Show

End Sub


Sub HA_ONB()

HA_Retour1.Show

End Sub

