VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GrootboekRoutecode 
   Caption         =   "Grootboek Routecode opvragen"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   OleObjectBlob   =   "GrootboekRoutecode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GrootboekRoutecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub G_RVLGND_Click()

ONDERWERP = "GROOTBOEK: " & FN & " | " & BdNm

If RCINACT = True Then
ONDERWERP = ONDERWERP & " | RC INACTIEF"
AI = True
Call RCIEXCEL(DFA, EM, t, SubTxT, SN, FN, BN, RC)
End If

If GEENRC = True Then ONDERWERP = ONDERWERP & " | GEEN ROUTECODE"

If OVRG = "" Then GoTo Act1
OVRG = "<p style=font-size:14px;font-family:corbel;color:black><b>Overig, toelichting; </b><i>" & OVRG & "</i>"
Act1:

If GEENRC = True Then RC = ""
If GEENRC = True Then GoTo Act2
If RC = "" Then GoTo Act2
RC = "<p style=font-size:14px;font-family:corbel;color:black><b>De volgende routecode is inactief: <i>" & RC & "</i></b>"

Act2:
GrootboekRoutecode.Hide

Call GB01_Grootboek(ONDERWERP, RC, EM, OVRG, DSPLEML, Title, Report, t)

End Sub
Sub GB01_Grootboek(ONDERWERP, RC, EM, OVRG, DSPLEML, Title, Report, t)

    Dim i As String
    Dim b As String
    Dim pthSig As String
    Dim strSig As String
    Dim pthFOOT As String
    Dim strFOOT As String
    Dim pthBREAK As String
    Dim strBREAK As String
    Dim strUser As String
    
    
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)

pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\GBSCRPTS\GB 01.htm"
If RCINACT = True Then pthSig = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\GBSCRPTS\GB 02.htm"
pthFOOT = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\GBSCRPTS\GBFOOT.htm"

strSig = FSO.OpenTextFile(pthSig).ReadAll
strFOOT = FSO.OpenTextFile(pthFOOT).ReadAll

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

    Dim fwd As Outlook.MailItem
    Dim itm As Object
         
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set fwd = itm.Forward
        
        Do Until fwd.Attachments.Count = 0
            fwd.Attachments.Remove (1)
        Loop


            CopyAttachments itm, fwd

fwd.SentOnBehalfOfName = "facturen@amsterdam.nl"
fwd.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
fwd.Recipients.Add "Grootboek1@amsterdam.nl"

fwd.Subject = ONDERWERP

fwd.HTMLBody = strSig & RC & OVRG & strFOOT & fwd.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK

fwd.DeferredDeliveryTime = DateAdd("s", 25, Now)

If DSPLEML.Value = False Then fwd.Send
 If DSPLEML.Value = True Then fwd.Display

    End If
 
    Set fwd = Nothing
    Set itm = Nothing
    
    If CbAFGH.Value = False Then Afgehandeld
    
    KNOP7
   

    End Sub
Sub RCIEXCEL(DFA, EM, t, SubTxT, SN, FN, BN, RC)

 Dim xlApp As Object
 Dim xlWB As Object
 Dim xlSheet As Object
 Dim rCount As Long
 Dim bXStarted As Boolean
 Dim enviro As String
 Dim strPath As String
 Dim Status As String
 
 Dim strColA, strColB, strColC, strColD, strColE, strColF, strColG, strColH As String
 
 'If AI = True Then
 Status = "INACTIEF"
               
     On Error Resume Next
     Set xlApp = GetObject(, "Excel.Application")
     If Err <> 0 Then
         Application.StatusBar = "Please wait while Excel source is opened ... "
         Set xlApp = CreateObject("Excel.Application")
         bXStarted = True
     End If
     On Error GoTo 0
     
xlApp.ScreenUpdating = False

Set xlWB = xlApp.Workbooks.Open("G:\FIN\11DebCred\accounthouders\Routecode Inactief\INACTIEVE ROUTECODES.xlsx")
Set xlSheet = xlWB.Sheets("INACT RC's")
'
'  xlSheet.Range("A1") = "Behandeld"
'  xlSheet.Range("B1") = "Recieved Time"
'  xlSheet.Range("C1") = "Sender"
'  xlSheet.Range("D1") = "Sender address"
'  xlSheet.Range("E1") = "Subject"
'  xlSheet.Range("F1") = "Factuurnummer"
'  xlSheet.Range("G1") = "Routecode"
'
rCount = xlSheet.Range("B" & xlSheet.Rows.Count).End(-4162).Row
rCount = rCount + 1

  xlSheet.Range("A" & rCount) = (Format(Now, "dd-mm-yyyy hh:mm:ss"))
  xlSheet.Range("B" & rCount) = t
  xlSheet.Range("C" & rCount) = SN
  xlSheet.Range("D" & rCount) = EM
  xlSheet.Range("E" & rCount) = SubTxT
  xlSheet.Range("F" & rCount) = FN
  xlSheet.Range("G" & rCount) = Status
  xlSheet.Range("H" & rCount) = RC

xlApp.ScreenUpdating = True

xlWB.Save '("H:\Mijn Documenten\merge\pdf\12345.xlsx")
xlWB.Close False
          
MsgBox "Done!"
            
End Sub

