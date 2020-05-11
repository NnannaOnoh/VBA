Attribute VB_Name = "HAREPLY"
Sub HA01_ReedsBetaald()

pthSig1 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox herinneringen & Aanmaningen\H&A01-1.htm"
pthSig2 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox herinneringen & Aanmaningen\H&A01-2.htm"
ONDERWERP = "Factuur reeds betaald | "

Call EXPORTGET(CREDITEUR, FACTUURNUMMER, FACTUURDATUM, BEDRAG, ROUTECODE, GBDTM, BHPAVMAIL)

Call HA_Bericht(CREDITEUR, RETDTM, FACTUURNUMMER, FACTUURDATUM, BEDRAG, ROUTECODE, GBDTM, BHPAVMAIL, ONDERWERP, pthSig1, pthSig2)

End Sub
Sub HA02_BijPAV()

pthSig1 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox herinneringen & Aanmaningen\H&A02-1.htm"
pthSig2 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox herinneringen & Aanmaningen\H&A02-2.htm"
ONDERWERP = "Factuur ter goedkeuring bij PAV/Budgethouder | "

Call EXPORTGET(CREDITEUR, FACTUURNUMMER, FACTUURDATUM, BEDRAG, ROUTECODE, GBDTM, BHPAVMAIL)

Call HA_Bericht(CREDITEUR, RETDTM, FACTUURNUMMER, FACTUURDATUM, BEDRAG, ROUTECODE, GBDTM, BHPAVMAIL, ONDERWERP, pthSig1, pthSig2)

End Sub
Sub HA03_InOmloop()

pthSig1 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox herinneringen & Aanmaningen\H&A03-1.htm"
pthSig2 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox herinneringen & Aanmaningen\H&A03-2.htm"
ONDERWERP = "Factuur in  omloop | "

Call EXPORTGET(CREDITEUR, FACTUURNUMMER, FACTUURDATUM, BEDRAG, ROUTECODE, GBDTM, BHPAVMAIL)

Call HA_Bericht(CREDITEUR, RETDTM, FACTUURNUMMER, FACTUURDATUM, BEDRAG, ROUTECODE, GBDTM, BHPAVMAIL, ONDERWERP, pthSig1, pthSig2)

End Sub
Sub HA04_Factuuronbekend()

pthSig1 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox herinneringen & Aanmaningen\H&A04-1.htm"
pthSig2 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox herinneringen & Aanmaningen\H&A04-2.htm"
ONDERWERP = "Factuur onbekend | "

       CREDITEUR = InputBox("Leverancier")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
           
           FACTUURNUMMER = InputBox("Factuurnummer")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
            FACTUURDATUM = InputBox("Factuurdatum")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select

            BEDRAG = InputBox("Bedrag")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
'Call EXPORTGET(CREDITEUR, FACTUURNUMMER, FACTUURDATUM, BEDRAG, ROUTECODE, GBDTM, BHPAVMAIL)

Call HA_Bericht(CREDITEUR, RETDTM, FACTUURNUMMER, FACTUURDATUM, BEDRAG, ROUTECODE, GBDTM, BHPAVMAIL, ONDERWERP, pthSig1, pthSig2)

End Sub
Sub HA05_FactuurRetour()

pthSig1 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox herinneringen & Aanmaningen\H&A05-1.htm"
pthSig2 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox herinneringen & Aanmaningen\H&A05-2.htm"
ONDERWERP = "Factuur retour verzonden | "

       RETDTM = InputBox("Datum Retour")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
    CREDITEUR = InputBox("Leverancier")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
           
           FACTUURNUMMER = InputBox("Factuurnummer")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
            FACTUURDATUM = InputBox("Factuurdatum")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select

            BEDRAG = InputBox("Bedrag")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
RETDTM = "Op " & RETDTM & " hebben wij de factuur naar u retour gestuurd en u geïnformeerd welke gegevens misten. "
    
'Call EXPORTGET(CREDITEUR, FACTUURNUMMER, FACTUURDATUM, BEDRAG, ROUTECODE, GBDTM, BHPAVMAIL)

Call HA_Bericht(CREDITEUR, RETDTM, FACTUURNUMMER, FACTUURDATUM, BEDRAG, ROUTECODE, GBDTM, BHPAVMAIL, ONDERWERP, pthSig1, pthSig2)

End Sub
Sub EXPORTGET(CREDITEUR, FACTUURNUMMER, FACTUURDATUM, BEDRAG, ROUTECODE, GBDTM, BHPAVMAIL)

Cells.Find(What:="Crediteur").Activate

Selection.Offset(1, 0).Select

CREDITEUR = Selection.Value


Cells.Find(What:="GB- datum").Activate

Selection.Offset(1, 0).Select

GBDTM = Selection.Value


Cells.Find(What:="Factuur- nummer").Activate

Selection.Offset(1, 0).Select

FACTUURNUMMER = Selection.Value


Cells.Find(What:="Factuur- datum").Activate

Selection.Offset(1, 0).Select

FACTUURDATUM = Selection.Value


Cells.Find(What:="Bruto- bedrag").Activate

Selection.Offset(1, 0).Select

BEDRAG = Selection.Value


Cells.Find(What:="Routing Code").Activate

Selection.Offset(1, 0).Select

ROUTECODE = Selection.Value

strPath = "G:\FIN\11DebCred\Crediteuren\60. Team Input\TIJDELIJK\Contactlijst per route.xlsx"
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

            Cells.Find(What:=ROUTECODE).Activate

            Selection.Offset(0, 1).Select

    BHPAVMAIL = Selection.Value
            
            xlWb.Close False
            
On Error Resume Next

MsgBox CREDITEUR & " " & FACTUURNUMMER & " " & FACTUURDATUM & " " & BEDRAG & " " & ROUTECODE & " " & BHPAVMAIL, , "export"

End Sub
Sub HA_Bericht(CREDITEUR, RETDTM, FACTUURNUMMER, FACTUURDATUM, BEDRAG, ROUTECODE, GBDTM, BHPAVMAIL, ONDERWERP, pthSig1, pthSig2)

Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem
Dim oReply As Outlook.MailItem
Dim objSelection As Outlook.Selection
Dim itm As Object

Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

If GBDTM = "" Then GBDTM = "nog niet betaald"
If FACTUURDATUM = "" Then FACTUURDATUM = "niet bekend"
If BEDRAG = "" Then BEDRAG = "niet bekend"

strSig1 = FSO.OpenTextFile(pthSig1).ReadAll
strSig2 = FSO.OpenTextFile(pthSig2).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

Set objOL = CreateObject("Outlook.Application")

Set objSelection = objOL.ActiveExplorer.Selection

For Each objMsg In objSelection

    Set oReply = objMsg.Reply
    oReply.SentOnBehalfOfName = "facturen@amsterdam.nl"
    oReply.BCC = BHPAVMAIL
    oReply.Recipients.Item(1).Resolve
    oReply.Subject = ONDERWERP & FACTUURNUMMER & " | " & CREDITEUR
    
    oReply.HTMLBody = strSig1 & "<table style=""width:50%; border: 1px solid black; text-align: center; border-collapse: collapse;""><tr><th=>FACTUURDATUM</th><th>FACTUURNUMMER</th><th>BEDRAG</th><th>DATUM BETALING</th>" _
                       & "</tr><tr><td>" & FACTUURDATUM & "</td><td>" & FACTUURNUMMER & "</td><td>" & BEDRAG & "</input></td><td><span style='background:yellow;mso-highlight:yellow'><i>" & GBDTM & "</i></span></td></tr></td></table>" _
            & strBREAK & RETDTM & strSig2 & oReply.HTMLBody & strBREAK


    MsgBox ROUTECODE, , "Naam goedkeurder"
                        
    oReply.Display

Next

Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
Set oReply = Nothing

'ActiveWorkbook.Close SaveChanges:=False
'Workbooks("H&A.XLSM").Close SaveChanges:=False

End Sub



