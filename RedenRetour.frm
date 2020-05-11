VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RedenRetour 
   Caption         =   "Reden Retour"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   OleObjectBlob   =   "RedenRetour.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RedenRetour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RC_Click()

r = "Terug AE Routecode"

eis = "AE"

RedenRetour.Hide

Call FAC09_Creditnotas(r, eis)

End Sub
Private Sub PDF_Click()

r = "Terug AE PDF format"

eis = "AE"

RedenRetour.Hide

Call FAC09_Creditnotas(r, eis)

End Sub
Private Sub IBAN_Click()

r = "Terug AE IBAN nummer"

eis = "AE"

RedenRetour.Hide

Call FAC09_Creditnotas(r, eis)

End Sub
Private Sub DEB_Click()

r = "Terug WE Debet vermelding"

eis = "WE"

RedenRetour.Hide

Call FAC09_Creditnotas(r, eis)

End Sub
Private Sub NAWLEV_Click()

r = "Terug WE N.A.W. Lev"

eis = "WE"

RedenRetour.Hide

Call FAC09_Creditnotas(r, eis)

End Sub
Private Sub NAWGA_Click()

r = "Terug WE N.A.W. Gemeente Amsterdam"

eis = "WE"

RedenRetour.Hide

Call FAC09_Creditnotas(r, eis)

End Sub
Private Sub FCTNR_Click()

r = "Terug WE Factuurnummer"

eis = "WE"

RedenRetour.Hide

Call FAC09_Creditnotas(r, eis)

End Sub
Private Sub FCTDTM_Click()

r = "Terug WE Factuurdatum"

eis = "WE"

RedenRetour.Hide

Call FAC09_Creditnotas(r, eis)

End Sub
Private Sub KVK_Click()

r = "Terug WE KvK"

eis = "WE"

RedenRetour.Hide

Call FAC09_Creditnotas(r, eis)

End Sub
Private Sub BTW_Click()

r = "Terug WE BTW nummer"

eis = "WE"

RedenRetour.Hide

Call FAC09_Creditnotas(r, eis)

End Sub
Private Sub BDRG_Click()

r = "Terug WE Brutobedrag"

eis = "WE"

RedenRetour.Hide

Call FAC09_Creditnotas(r, eis)

End Sub
Private Sub DNB_Click()

r = "Debet niet bekend"

RedenRetour.Hide

Call DBNBERICHT(r)

End Sub

Public Sub FAC09_Creditnotas(r, eis)
    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As attachment
    Dim Report As String
    Dim t As Date
    
      
    For Each selItem In Application.ActiveExplorer.Selection
        If selItem.Class = olMail Then
            Set aMail = selItem
            For Each aAttach In aMail.Attachments
                Report = Report & GetAttachmentInfo(aAttach)
                Report = Report & ", "
            Next
            
             t = selItem.ReceivedTime
            em = selItem.SenderEmailAddress
            
            Call FAC09_Creditnotas2("", Report, r, t, em, eis)
            
        End If
    Next
End Sub

Sub FAC09_Creditnotas2(Title As String, Report As String, r, t, em, eis)

    Dim i As String
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
    
    Dim rCount As String
    
    Dim strColA, strColB, strColC, strColD, strColE, strColF, strColG, strColH, strColI, strColJ, strColK, strColL, strColM, strColN, strColO, strColP, strColQ, strColR As String
       
    
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)

pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 03.htm"
strSig = FSO.OpenTextFile(pthSig).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

strPath = "G:\FIN\11DebCred\Crediteuren\20. Verwerking facturen\231. Creditfacturen\Teruggestuurde Creditnotas.xlsx"

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
    
            Fd = InputBox("Factuurdatum")
    Select Case StrPtr(Fd)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
            Fl = InputBox("Factuurbedrag")
    Select Case StrPtr(Fl)
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

rpl.Subject = "Teruggestuurd/" & i & "/" & b & "/CR; " & eis

rpl.HTMLBody = strSig & rpl.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK
rpl.Display
rpl.DeferredDeliveryTime = DateAdd("s", 25, Now)

    End If

   
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
rCount = xlSheet.Range("D" & xlSheet.Rows.Count).End(-4162).Row
rCount = rCount + 1

    'strColA = ""
    strColB = strUser
    strColC = t
    strColD = (Format(Now, "dd-mm-yyyy hh:mm:ss"))
    strColE = b
    strColF = em
    strColG = Fd
    strColH = i
    strColI = Fl
    strColJ = r
    'strColK = ""
    'strColL = ""
    'strColM = ""
    strColN = "Open"
    'strColO = ""
    'strColP = ""
    strColQ = Report
    'strColR = ""

  'xlSheet.Range("A" & rCount) = strColA
  xlSheet.Range("B" & rCount) = strColB
  xlSheet.Range("C" & rCount) = strColC
  xlSheet.Range("D" & rCount) = strColD
  xlSheet.Range("E" & rCount) = strColE
  xlSheet.Range("F" & rCount) = strColF
  xlSheet.Range("G" & rCount) = strColG
  xlSheet.Range("H" & rCount) = strColH
  xlSheet.Range("I" & rCount).Style = "Currency"
  xlSheet.Range("I" & rCount) = strColI
  xlSheet.Range("J" & rCount) = strColJ
  'xlSheet.Range("K" & rCount) = strColK
  'xlSheet.Range("L" & rCount) = strColL
  'xlSheet.Range("M" & rCount) = strColM
  xlSheet.Range("N" & rCount) = strColN
  'xlSheet.Range("O" & rCount) = strColO
  'xlSheet.Range("P" & rCount) = strColP
  xlSheet.Range("Q" & rCount) = strColQ
  'xlSheet.Range("R" & rCount) = strColR

  
  rCount = rCount + 1
 
Application.DisplayAlerts = False
 
xlWb.Save '("G:\FIN\11DebCred\Crediteuren\20. Verwerking facturen\231. Creditfacturen\Teruggestuurde Creditnotas.xlsx")
xlWb.Close False
Application.DisplayAlerts = False


     'If bXStarted Then
     '    xlApp.Quit
     'End If

     Set xlApp = Nothing
     Set xlWb = Nothing
     Set xlSheet = Nothing
     Set rpl = Nothing
     Set itm = Nothing

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

Public Sub DBNBERICHT(r)
    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As attachment
    Dim Report As String
    Dim t As Date
    
      
    For Each selItem In Application.ActiveExplorer.Selection
        If selItem.Class = olMail Then
            Set aMail = selItem
            For Each aAttach In aMail.Attachments
                Report = Report & GetAttachmentInfo(aAttach)
                Report = Report & ", "
            Next
            
             t = selItem.ReceivedTime
            em = selItem.SenderEmailAddress
            
BN = InputBox("Bestandsnaam", Default, Report)
    Select Case StrPtr(BN)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
            
            Call HANDMATIG(Report, BN)
            
            Call DBNBERICHT2("", Report, r, t, em, BN)
            
        End If
    Next
End Sub
Sub DBNBERICHT2(Title As String, Report As String, r, t, em, BN)


    Dim fwd As Outlook.MailItem
    Dim itm As Object
    
    Dim i As String
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
    
    Dim rCount As String
    
    Dim strColA, strColB, strColC, strColD, strColE, strColF, strColG, strColH, strColI, strColJ, strColK, strColL, strColM, strColN, strColO, strColP, strColQ, strColR As String
       
    
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)

pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 03.htm"
strSig = FSO.OpenTextFile(pthSig).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

strPath = "G:\FIN\11DebCred\Crediteuren\20. Verwerking facturen\231. Creditfacturen\Teruggestuurde Creditnotas.xlsx"

        
    'Set itm = GetCurrentItem()
    'If Not itm Is Nothing Then
    '    Set fwd = itm.Forward
    '    CopyAttachments itm, fwd
        
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
    
            Fd = InputBox("Factuurdatum")
    Select Case StrPtr(Fd)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
            Fl = InputBox("Factuurbedrag")
    Select Case StrPtr(Fl)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
      
    strUser = Left(Environ("USERNAME"), 3)
   
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set fwd = itm.Forward
        
        Do Until fwd.Attachments.Count = 0
            fwd.Attachments.Remove (1)
        Loop
        
        fwd.SentOnBehalfOfName = "facturen@amsterdam.nl"
        fwd.Recipients.Add "srvc18VR@amsterdam.nl"
        
        'CopyAttachments itm, fwd
        fwd.Attachments.Add "H:\Mijn Documenten\merge\pdf\OLAttachments\watermerk\" & BN & ".pdf"
        fwd.Subject = "Debet niet bekend/" & i & "/" & b & "/CR; "

        fwd.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
        fwd.HTMLBody = fwd.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser
        
        fwd.DeferredDeliveryTime = DateAdd("s", 25, Now)
        fwd.Display
        
        
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
rCount = xlSheet.Range("D" & xlSheet.Rows.Count).End(-4162).Row
rCount = rCount + 1

    'strColA = ""
    strColB = strUser
    strColC = t
    strColD = (Format(Now, "dd-mm-yyyy hh:mm:ss"))
    strColE = b
    strColF = em
    strColG = Fd
    strColH = i
    strColI = Fl
    strColJ = r
    'strColK = ""
    'strColL = ""
    'strColM = ""
    strColN = "Open"
    'strColO = ""
    'strColP = ""
    strColQ = Report
    'strColR = ""

  'xlSheet.Range("A" & rCount) = strColA
  xlSheet.Range("B" & rCount) = strColB
  xlSheet.Range("C" & rCount) = strColC
  xlSheet.Range("D" & rCount) = strColD
  xlSheet.Range("E" & rCount) = strColE
  xlSheet.Range("F" & rCount) = strColF
  xlSheet.Range("G" & rCount) = strColG
  xlSheet.Range("H" & rCount) = strColH
  xlSheet.Range("I" & rCount).Style = "Currency"
  xlSheet.Range("I" & rCount) = strColI
  xlSheet.Range("J" & rCount) = strColJ
  'xlSheet.Range("K" & rCount) = strColK
  'xlSheet.Range("L" & rCount) = strColL
  'xlSheet.Range("M" & rCount) = strColM
  xlSheet.Range("N" & rCount) = strColN
  'xlSheet.Range("O" & rCount) = strColO
  'xlSheet.Range("P" & rCount) = strColP
  xlSheet.Range("Q" & rCount) = strColQ
  'xlSheet.Range("R" & rCount) = strColR

  
  rCount = rCount + 1
 
Application.DisplayAlerts = False
 
xlWb.Save '("G:\FIN\11DebCred\Crediteuren\20. Verwerking facturen\231. Creditfacturen\Teruggestuurde Creditnotas.xlsx")
xlWb.Close False
Application.DisplayAlerts = False


     'If bXStarted Then
     '    xlApp.Quit
     'End If

     Set xlApp = Nothing
     Set xlWb = Nothing
     Set xlSheet = Nothing
     Set rpl = Nothing
     Set itm = Nothing

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

End If

End Sub



