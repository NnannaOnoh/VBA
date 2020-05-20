Attribute VB_Name = "Module2"
Public Function AAPJEPUNTJE2(BdNm, INTERN, SN, SubTxT, EM, t, WaardenEM, WaardenBdNm)

    Dim selItem As Object
    Dim aMail As MailItem
    
    Dim AAPJE, PUNTJE As Integer
    Dim TEMPStr, PUNTJE2 As String
    
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

On Error Resume Next
      pthEM = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\Exceptions\DONOTBOUNCE.txt"
      strEM = FSO.OpenTextFile(pthEM).ReadAll
     pthEM1 = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\Exceptions\NOREPLY.txt"
     strEM1 = FSO.OpenTextFile(pthEM1).ReadAll
    pthBdNM = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\Exceptions\FINANCEPF.txt"
    strBdNM = FSO.OpenTextFile(pthBdNM).ReadAll
   
    For Each selItem In Application.ActiveExplorer.Selection
        If selItem.Class = olMail Then
            Set aMail = selItem
            
    AM = "/O=MAILORG/OU=EXCHANGE ADMINISTRATIVE GROUP"
SubTxT = aMail.Subject                                              '<- SUBJECT | E-MAIL ONDERWERP
    EM = aMail.SenderEmailAddress                                   '<- EMAIL | E-MAIL ADRES AFZENDER
    SN = aMail.SenderName                                           '<- SENDERNAME | E-MAIL DISPLAYNAAM AFZENDER
     t = aMail.ReceivedTime                                          '<- TIJD | E-MAIL ONTVANGSTDATUM & TIJD
    
    INTERN = InStr(1, EM, AM, vbTextCompare)
   End If
  
If INTERN > 0 Then
BdNm = SN
Exit Function
Else
AAPJE = InStr(1, EM, "@")
TEMPStr = Right(EM, Len(EM) - AAPJE + 0)
End If

EM1 = Left(EM, AAPJE - 1)

On Error Resume Next
PUNTJE2 = Split(TEMPStr, ".")(2)
PUNTJE = InStr(1, TEMPStr, ".")

If Not PUNTJE2 = "" Then
BdNm = UCase(Left(TEMPStr, Len(TEMPStr) - (Len(PUNTJE2) + 1)))
Else
BdNm = UCase(Left(TEMPStr, PUNTJE - 1))                             '<- BEDRIJFSNAAM | E-MAIL EXTRACT
End If

AAPJEPUNTJE = BdNm

        Exceptions = strBdNM   '<------ finance processing firm
        If InStr(1, Exceptions, "|" & BdNm & "|", vbTextCompare) Then WaardenBdNm = True
        
        Exceptions = strEM   '<------ bounce e-mail address
        If InStr(1, Exceptions, "|" & EM & "|", vbTextCompare) Then WaardenEM = True
        
        Exceptions = strEM1   '<------ noreply
        If InStr(1, Exceptions, "|" & EM1 & "|", vbTextCompare) Then WaardenEM = True

Next


End Function
