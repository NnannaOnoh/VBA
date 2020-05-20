VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Waarden 
   Caption         =   "Factuur"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   OleObjectBlob   =   "Waarden.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Waarden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub F_CVLGND_Click()

Waarden.Hide
        
   If CbHA.Value = True Then
            Call Herinnering
                    Exit Sub
                      End If

   If CbGB.Value = True Then
 Call Grootboek01(FN, RC, BdNm, EM, SN, SubTxT, CbAFGH, DSPLEML, t, OVRG)
                    Exit Sub
                      End If
   If CbIO = True And _
    CbAEAD = True And _
      CbFN = True And _
      CbFD = True And _
    CbWEAD = True And _
    CbWEBT = True And _
    CbBTNR = True And _
    CbIBAN = True And _
    CbBTNR = True And _
    CbWEKV = True And _
     CbPDF = True And _
  CbANDERS = False Then

  If DFA2 = True Then DFA = "srvc18VR@amsterdam.nl"
  If DFA3 = True Then DFA = "srvc47ACAM@amsterdam.nl"
  If DFA4 = True Then DFA = "srvc90SR@amsterdam.nl"
  
    If DFA = "" Then
    MsgBox "Vul DFA in", vbOKOnly, "DFA"
    Waarden.Show
    Exit Sub
    End If
    
   If CbARC.Value = True And RC = "" Then
   MsgBox "Vul Routecode in"
                   Exit Sub
                     End If
    
   If CbARC.Value = True And CbMrg.Value = True Then
           Call MERGE(DFA, PDF, FN, CbMrg, CbAFGH, DSPLEML, SNDEML, CbARC, CbINK, EM, t, SubTxT, SN, REC, RC)
                    Exit Sub
                      End If
                      
    If CbINK.Value = True And CbMrg.Value = True Then
           Call MERGE(DFA, PDF, FN, CbMrg, CbAFGH, DSPLEML, SNDEML, CbARC, CbINK, EM, t, SubTxT, SN, REC, RC)
                     Exit Sub
                       End If

   If CbARC.Value = True Then
           Call RCALTERNATIEF(DFA, PDF, DestFile, FN, CbAFGH, DSPLEML, SNDEML, CbARC, CbMrg, EM, t, SubTxT, SN, REC, RC)
                     Exit Sub
                       End If

    If CbINK.Value = True Then
           Call IOINACTIEF(DFA, PDF, DestFile, FN, CbAFGH, DSPLEML, SNDEML, CbARC, CbINK, CbMrg, EM, t, SubTxT, SN, REC, RC)
                     Exit Sub
                       End If

    If CbMrg.Value = True Then
           Call MERGE(DFA, PDF, FN, CbMrg, CbAFGH, DSPLEML, SNDEML, CbARC, CbINK, EM, t, SubTxT, SN, REC, RC)
                     Exit Sub
                       End If

Call Factuur_Compleet(PDF, DFA, FN, DSPLEML, CbARC, CbINK, CbAFGH, CbSendAll, CbAB, EM, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

Exit Sub
                     
                     
Else

CbAE = False
CbWE = False
CbFM = False
 
If CbONb = True Then Waarden.FN = "ONBEKEND"

If CbAEAD = False Then AEAD = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\AEAD.htm"
If CbAEAD = False Then CbAE = True

If CbIO = False Then AEIO = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\AEIO1.htm"
If CbIO = False Then AEEO = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\AEEO.htm"
If CbIO = False Then AERC = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\AERC.htm"
If CbIO = False Then CbAE = True

'If CbAERC = False Then CbAE = True
'If CbIO = False And CbAERC = True Then AEIO = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\AEIO1.htm"

If CbPDF = False Then FMPD = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\FMPD.htm"
If CbPDF = False Then CbFM = True

If CbWEAD = False Then WEAD = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEAD.htm"
If CbWEAD = False Then CbWE = True

If CbWEBT = False Then BTNR = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\BTNR.htm" '<--- ongelijke waarde
If CbWEBT = False Then CbWE = True
If CbBTNR = False Then WEBT = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEBT.htm" '<--- ongelijke waarde
If CbBTNR = False Then CbWE = True
If CbFD = False Then WEFD = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEFD.htm"
If CbFD = False Then CbWE = True
If CbFN = False Then WEFN = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEFN.htm"
If CbFN = False Then CbWE = True
If CbIBAN = False Then WEIB = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEIB.htm"
If CbIBAN = False Then CbWE = True
If CbWEKV = False Then WEKV = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEKV.htm"
If CbWEKV = False Then CbWE = True

If CbWE = True Then WEBD = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\WEBD.htm"

If CbAE = False Then AE = ""
If CbAE = True Then AE = "AE"
If CbAE = True And CbWE = True Then AE = "AE; "
If CbAE = True And CbFM = True Then AE = "AE; "

If CbWE = False Then WE = ""
If CbWE = True Then WE = "WE"
If CbWE = True And CbFM = True Then WE = "WE; "


If CbFM = False Then FM = ""
If CbFM = True Then FM = "FM"

If CbONb = True Then FN = "ONBEKEND"

REDEN = AE & WE & FM

If Not OVRG = "" Then OVRG = "<p style=font-size:14px;font-family:corbel;color:black><b>Overig, toelichting; </b><i>" & OVRG & "</i>"

'If CbRC = True Then OVRG = "routecode " & OVRG & " is niet valide"

Call Factuur_Retour(CbAB, PDF, FN, BdNm, CbMrg, CbEM, EM, ONDERWERP, AEAD, AEIO, AERC, AEEO, FMPD, WEAD, WEBT, BTNR, WEFD, WEFN, WEIB, WEKV, WEBD, OVRG, REDEN, DSPLEML, CbAFGH, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)
    
End If
        
End Sub
Private Sub EM_Change()
CbEM = True
Waarden.EM.BackColor = vbWhite
End Sub
Private Sub BdNm_Change()
Waarden.BdNm.BackColor = vbWhite
End Sub
Private Sub CbONb_Change()
If CbONb = True Then
Waarden.FN = "ONBEKEND"
Else
Waarden.FN = FN1
End If
End Sub
Private Sub COMPLEET_Click()

Waarden.Hide

If DFA2 = True Then DFA = "srvc18VR@amsterdam.nl"
If DFA3 = True Then DFA = "srvc47ACAM@amsterdam.nl"
If DFA4 = True Then DFA = "srvc90SR@amsterdam.nl"

Call Factuur_Compleet(PDF, DFA, FN, DSPLEML, CbARC, CbINK, CbAFGH, CbSendAll, CbAB, EM, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub Close1_Click()
Me.Caption = "PDF bestanden uit map verwijderen....."
KillAll
Me.Caption = "Waarden"
Waarden.Hide

End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 1 Then Cancel = 1
End Sub
