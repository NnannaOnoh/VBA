VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FacturenRetour 
   Caption         =   "Factuur Retour"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   OleObjectBlob   =   "FacturenRetour.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FacturenRetour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub F_RVLGND_Click()

If CbAEAD = True Then AEAD = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\AEAD.htm"
If CbAEAD = True Then CbAE = True
If CbAEIO = True Then AEIO = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\AEIO.htm"
If CbAEIO = True Then CbAE = True
If CbAERC = True Then AERC = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\AERC.htm"
If CbAERC = True Then CbAE = True
If CbAEIO = True And CbAERC = True Then AEEO = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\AEEO.htm"
If CbAEIO = True And CbAERC = True Then AEIO = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\AEIO1.htm"
If CbFMPD = True Then FMPD = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\FMPD.htm"
If CbFMPD = True Then CbFM = True
If CbWEAD = True Then WEAD = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEAD.htm"
If CbWEAD = True Then CbWE = True
If CbWEBT = True Then WEBT = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEBT.htm"
If CbWEBT = True Then CbWE = True
If CbWEFD = True Then WEFD = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEFD.htm"
If CbWEFD = True Then CbWE = True
If CbWEFN = True Then WEFN = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEFN.htm"
If CbWEFN = True Then CbWE = True
If CbWEIB = True Then WEIB = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEIB.htm"
If CbWEIB = True Then CbWE = True
If CbWEKV = True Then WEKV = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEKV.htm"
If CbWEKV = True Then CbWE = True

If CbWE = True Then WEBD = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEBD.htm"

If CbAE = False Then AE = ""
If CbAE = True Then AE = "AE"
If CbAE = True And CbWE = True Then AE = "AE; "
If CbAE = True And CbOR = True Then AE = "AE; "
If CbAE = True And CbFM = True Then AE = "AE; "
If CbAE = True And CbCR = True Then AE = "AE; "

If CbWE = False Then WE = ""
If CbWE = True Then WE = "WE"
If CbWE = True And CbOR = True Then WE = "WE; "
If CbWE = True And CbFM = True Then WE = "WE; "
If CbWE = True And CbCR = True Then WE = "WE; "

If CbOR = False Then OuR = ""
If CbOR = True Then OuR = "OR"
If CbOR = True And CbFM = True Then OuR = "OR; "
If CbOR = True And CbCR = True Then OuR = "OR; "

If CbFM = False Then FM = ""
If CbFM = True Then FM = "FM"
If CbFM = True And CbCR = True Then FM = "FM; "

If CbCR = False Then CR = ""
If CbCR = True Then CR = "CR"

If CbFN = True Then FN = "onbekend"

ONDERWERP = "Teruggestuurd/" & Trim(FN) & "/" & Trim(BN) & "/" & AE & WE & OuR & FM & CR

If OVRG = "" Then GoTo Act1

If CbRC = True Then OVRG = "routecode " & OVRG & " is niet valide"

OVRG = "<p style=font-size:14px;font-family:corbel;color:black><b>Overig, toelichting; </b><i>" & OVRG & "</i>"

Act1:

FacturenRetour.Hide

Call showmsg2(ONDERWERP, AEAD, AEIO, AERC, AEEO, CbFM, FMPD, WEAD, WEBT, WEFD, WEFN, WEIB, WEKV, WEBD, CbEM, EM, OVRG, DSPLEML, Title, Report)

End Sub
Sub showmsg2(ONDERWERP, AEAD, AEIO, AERC, AEEO, CbFM, FMPD, WEAD, WEBT, WEFD, WEFN, WEIB, WEKV, WEBD, CbEM, EM, OVRG, DSPLEML, Title, Report)
    
    Dim pthBREAK As String
    Dim strBREAK As String
    Dim strUser As String
    
    Dim HEAD As String

Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

HEAD = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\HEAD.htm"
strHEAD = FSO.OpenTextFile(HEAD).ReadAll

On Error Resume Next
strAEAD = FSO.OpenTextFile(AEAD).ReadAll
strAEIO = FSO.OpenTextFile(AEIO).ReadAll
strAERC = FSO.OpenTextFile(AERC).ReadAll
strAEEO = FSO.OpenTextFile(AEEO).ReadAll
strFMPD = FSO.OpenTextFile(FMPD).ReadAll
strWEAD = FSO.OpenTextFile(WEAD).ReadAll
strWEBT = FSO.OpenTextFile(WEBT).ReadAll
strWEFD = FSO.OpenTextFile(WEFD).ReadAll
strWEFN = FSO.OpenTextFile(WEFN).ReadAll
strWEIB = FSO.OpenTextFile(WEIB).ReadAll
strWEKV = FSO.OpenTextFile(WEKV).ReadAll
strWEBD = FSO.OpenTextFile(WEBD).ReadAll

FOOT = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\FOOT.htm"
strFOOT = FSO.OpenTextFile(FOOT).ReadAll

    Dim rpl As Outlook.MailItem
    Dim itm As Object
    Dim verzonden As Integer
         
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        If CbEM = False Then Set rpl = itm.Reply
        If CbEM = True Then Set rpl = CreateItem(0)
            If CbFM = True Then CopyAttachmentsFM itm, rpl
            If CbFM = False Then CopyAttachments itm, rpl

If CbEM.Value = True Then rpl.Recipients.Add EM

rpl.SentOnBehalfOfName = "facturen@amsterdam.nl"
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\archief\flyer_1_administratie_a5_digitaal.pdf"
       
rpl.Subject = ONDERWERP

rpl.HTMLBody = "<p style=font-size:14px;font-family:corbel;color:black>" _
                      & strHEAD & strAERC & strAEEO & strAEIO & strAEAD & strWEAD & strWEBT & strWEFD & strWEFN & strWEIB & strWEKV & strFMPD _
                      & strBREAK & OVRG & strWEBD & strFOOT & strBREAK & rpl.HTMLBody _
                      & "<p class=MsoNormal><o:p> </o:p></P><div><div style='border:none;border-top:solid #B5C4DF 1.0pt;padding:3.0pt 0cm 0cm 0cm'>" _
                      & "<p class=MsoNormal style='line-height:normal'><br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<p style=font-size:14px;font-family:corbel;color:white><br>" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK & "</span>"
                      
rpl.DeferredDeliveryTime = DateAdd("s", 25, Now)
                      
If ObDSPLEML.Value = False Then rpl.Send
 If ObDSPLEML.Value = True Then rpl.Display

    End If
 
    Set rpl = Nothing
    Set itm = Nothing
    
     Call KNOP2
    
If CbRTR.Value = False Then Retour

    End Sub
