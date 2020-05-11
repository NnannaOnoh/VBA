Attribute VB_Name = "HASCRIPTS"
Sub HA01_Factuur_Reeds_Betaald()

Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

pthSig1 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox herinneringen & Aanmaningen\H&A01-1.htm"
strSig1 = FSO.OpenTextFile(pthSig1).ReadAll
pthSig2 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox herinneringen & Aanmaningen\H&A01-2.htm"
strSig2 = FSO.OpenTextFile(pthSig2).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

    Dim rpl As Outlook.MailItem
    Dim itm As Object
    Dim verzonden As Integer
    Dim DateClicked As Date
    
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
    
    d = InputBox("Datum betaling")
    Select Case StrPtr(d)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
            be = InputBox("Bedrag")
    Select Case StrPtr(be)
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
'rpl.Recipients.Add ""
rpl.Attachments.Add ("G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\archief\flyer_1_administratie_a5_digitaal.pdf")
rpl.Subject = "Factuurnummer " & i & " | " & b & " reeds betaald."

rpl.Display

rpl.HTMLBody = strSig1 & "<table style=""width:50%; border: 1px solid black; text-align: center; border-collapse: collapse;""><tr><th=>FACTUURNUMMER</th><th>BEDRAG</th><th>DATUM BETALING</th>" _
                       & "</tr><tr><td>" & i & "</td><td>" & be & "</td><td>" & d & "</td></tr></td></table>" & strSig2 & rpl.HTMLBody & strBREAK

    End If
 
    Set rpl = Nothing
    Set itm = Nothing
    
    'Call Retour

    End Sub
