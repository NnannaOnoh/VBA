Attribute VB_Name = "Mails"
Sub Factuur_Compleet(PDF, DFA, FN, DSPLEML, CbARC, CbINK, CbAFGH, CbSendAll, CbAB, EM, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)
  
Dim fwd As Outlook.MailItem
Dim itm As Object
Dim strUser As String

If CbARC = True Then MsgBox "Gegevens RC: " & RC & " ingevuld op PDF en Opgeslagen."

If CbINK = True Then MsgBox "Gegevens IO: INKOOPORDERNUMMER ingevuld op PDF en Opgeslagen."
    
strUser = Left(Environ("USERNAME"), 3)


PDF = "H:\Mijn Documenten\merge\pdf\OLAttachments\" & PDF
  
Set itm = GetCurrentItem()
If Not itm Is Nothing Then
Set fwd = itm.Forward
End If
           Do Until fwd.Attachments.Count = 0
                    fwd.Attachments.Remove (1)
           Loop
        
                    fwd.SentOnBehalfOfName = "Facturen@amsterdam.nl"
                    fwd.Recipients.Add DFA
                                
                If CbSendAll = True Or _
                CbAB = True Then
                    CopyAttachments itm, fwd
                    fwd.Subject = fwd.Subject & " " & FNABCMPLT
                Else
                    fwd.Attachments.Add PDF
                    fwd.Subject = fwd.Subject & " " & FN
                End If
        
                    fwd.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
                    fwd.HTMLBody = fwd.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & FN _
                    & FNABCMPLT & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                    & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser
        
                    fwd.DeferredDeliveryTime = DateAdd("s", 25, Now)

                If DSPLEML.Value = True Then
                    fwd.Display
                Else
                    fwd.Send
                End If

    Set fwd = Nothing
    Set itm = Nothing
    
KNOP1

If CbAFGH = False Then Afgehandeld

KillAll
    
End Sub
Sub Factuur_Retour(CbAB, PDF, FN, BdNm, CbMrg, CbEM, EM, ONDERWERP, AEAD, AEIO, AERC, AEEO, FMPD, WEAD, WEBT, BTNR, WEFD, WEFN, WEIB, WEKV, WEBD, OVRG, REDEN, DSPLEML, CbAFGH, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)
    
    Dim pthBREAK, strBREAK, strCbEM1, strCbEM2 As String
    Dim strUser As String
    
    Dim HEAD As String

Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

strUser = Left(Environ("USERNAME"), 3)
PDF = "H:\Mijn Documenten\merge\pdf\OLAttachments\" & PDF

pthBREAK = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\BREAKER.htm"
strBREAK = FSO.OpenTextFile(pthBREAK).ReadAll

HEAD = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\HEAD.htm"
strHEAD = FSO.OpenTextFile(HEAD).ReadAll

CbEM1 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 06.htm"
strCbEM1 = FSO.OpenTextFile(CbEM1).ReadAll

CbEM2 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FAC 05.htm"
strCbEM2 = FSO.OpenTextFile(CbEM2).ReadAll

On Error Resume Next
strAEAD = FSO.OpenTextFile(AEAD).ReadAll
strAEIO = FSO.OpenTextFile(AEIO).ReadAll
strAERC = FSO.OpenTextFile(AERC).ReadAll
strAEEO = FSO.OpenTextFile(AEEO).ReadAll
strFMPD = FSO.OpenTextFile(FMPD).ReadAll
strWEAD = FSO.OpenTextFile(WEAD).ReadAll
strBTNR = FSO.OpenTextFile(BTNR).ReadAll
strWEBT = FSO.OpenTextFile(WEBT).ReadAll
strWEFD = FSO.OpenTextFile(WEFD).ReadAll
strWEFN = FSO.OpenTextFile(WEFN).ReadAll
strWEIB = FSO.OpenTextFile(WEIB).ReadAll
strWEKV = FSO.OpenTextFile(WEKV).ReadAll
strWEBD = FSO.OpenTextFile(WEBD).ReadAll
On Error GoTo 0

FOOT = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\FACTSCRPTS\FOOT.htm"
strFOOT = FSO.OpenTextFile(FOOT).ReadAll

    Dim rpl As Outlook.MailItem
    Dim itm As Object
    Dim verzonden As Integer
         
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
   
nietaangepast:
      
    NOREPLY1 = "noreply"
  
    NOREPLY = InStr(1, EM, NOREPLY1, vbTextCompare)
    
    If NOREPLY > 0 Then
        
    CbEM.Value = True
        
    EM = InputBox("NOREPLY E-MAILADRES", "Geef E-mailadres aan", EM)
        
    End If
   
    NOREPLY = InStr(1, EM, NOREPLY1, vbTextCompare)
    
    If NOREPLY > 0 Then GoTo nietaangepast
        
            If CbEM = False Then
            Set rpl = itm.Reply
            Else
            Set rpl = CreateItem(0)
            End If
            
If CbEM.Value = True Then rpl.Recipients.Add EM

rpl.SentOnBehalfOfName = "facturen@amsterdam.nl"
rpl.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0

            If CbMrg = True Or _
            CbAB = True Or _
            Not FMPD = "" Then
            CopyAttachmentsFM itm, rpl
            ONDERWERP = "Teruggestuurd/" & FNABCMPLT & "/" & Trim(BdNm) & "/" & REDEN
            Else
            If Not PDF = "" Then
            rpl.Attachments.Add PDF
            ONDERWERP = "Teruggestuurd/" & FN & "/" & Trim(BdNm) & "/" & REDEN
            
            End If
                
            End If
'ONDERWERP = "Teruggestuurd/" & Trim(FN) & "/" & Trim(BdNm) & "/" & REDEN

rpl.Subject = ONDERWERP

If CbEM = 1 Then
rpl.HTMLBody = "<p style=font-size:14px;font-family:corbel;color:black>" _
                      & strCbEM1 & rpl.HTMLBody _
                      & "<p class=MsoNormal><o:p> </o:p></P><div><div style='border:none;border-top:solid #B5C4DF 1.0pt;padding:3.0pt 0cm 0cm 0cm'>" _
                      & "<p class=MsoNormal style='line-height:normal'><br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<p style=font-size:14px;font-family:corbel;color:white><br>" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK & "</span>"
ElseIf CbEM = 2 Then
rpl.HTMLBody = "<p style=font-size:14px;font-family:corbel;color:black>" _
                      & strCbEM2 & rpl.HTMLBody _
                      & "<p class=MsoNormal><o:p> </o:p></P><div><div style='border:none;border-top:solid #B5C4DF 1.0pt;padding:3.0pt 0cm 0cm 0cm'>" _
                      & "<p class=MsoNormal style='line-height:normal'><br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<p style=font-size:14px;font-family:corbel;color:white><br>" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK & "</span>"
Else
rpl.HTMLBody = "<p style=font-size:14px;font-family:corbel;color:black>" _
                      & strHEAD & strAERC & strAEEO & strAEIO & strAEAD & strWEAD & strBTNR & strWEBT & strWEFD & strWEFN & strWEIB & strWEKV & strFMPD _
                      & strBREAK & OVRG & strWEBD & strFOOT & strBREAK & rpl.HTMLBody _
                      & "<p class=MsoNormal><o:p> </o:p></P><div><div style='border:none;border-top:solid #B5C4DF 1.0pt;padding:3.0pt 0cm 0cm 0cm'>" _
                      & "<p class=MsoNormal style='line-height:normal'><br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<p style=font-size:14px;font-family:corbel;color:white><br>" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser & strBREAK & "</span>"
End If
                      
rpl.DeferredDeliveryTime = DateAdd("s", 25, Now)
                      
                If DSPLEML.Value = True Then
                   rpl.Display
                Else
                   rpl.Send
                End If
    End If
 
    Set rpl = Nothing
    Set itm = Nothing
    
     Call KNOP2
    
If CbAFGH.Value = False Then Retour

KillAll

    End Sub

