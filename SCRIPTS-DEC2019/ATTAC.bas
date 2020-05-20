Attribute VB_Name = "ATTAC"
Public Function SaveAtt(strFolderpath, PDFCount, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection

Dim i As Long
Dim a As Long
Dim lngCount As Long

Dim PDFName(1 To 25) As String

Dim strFile As String
Dim strDeletedFiles As String

    strFolderpath = ("H:\Mijn Documenten\merge\pdf")
    On Error Resume Next

    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    strFolderpath = strFolderpath & "\OLAttachments\"

    For Each objMsg In objSelection

    Set objAttachments = objMsg.Attachments
    lngCount = objAttachments.Count
        
    If lngCount > 0 Then
    
    For i = lngCount To 1 Step -1
    
    If UCase(Right(objAttachments.Item(i).FileName, 3)) = "PDF" Then
        
    strFile = objAttachments.Item(i).FileName
    
    strFile = Replace(strFile, ",", " ")
    
    
                            a = a + 1
               PDFName(a) = strFile                             '<- PDF NAAM | BESTANDSNAMEN VAN PDF BESTANDEN
                 PDFCount = PDFCount + 1                        '<- PDF TELLING | AANTAL PDF BESTANDEN IN E-MAIL
            
            
    strFile = strFolderpath & strFile
   
    objAttachments.Item(i).SaveAsFile strFile
    
    FNAB = Left(PDFName(a), Len(PDFName(a)) - 4) & " & " & FNAB
        
         Else
         
    strFile = objAttachments.Item(i).FileName
    
    FNAB2 = Left(strFile, Len(strFile) - 4) & " & " & FNAB2

        End If
              
    Next i
    End If
    
    Next
    
    FNAB = Left(FNAB, Len(FNAB) - 3)
    FNAB2 = Left(FNAB2, Len(FNAB2) - 3)
    
    FNABCMPLT = FNAB & FNAB2
    If Not FNAB = "" And Not FNAB2 = "" Then FNABCMPLT = FNAB & " / " & FNAB2
    
    'MsgBox FNABCMPLT
    
 TxTBN1 = PDFName(1)
 TxTBN2 = PDFName(2)
 TxTBN3 = PDFName(3)
 TxTBN4 = PDFName(4)
 TxTBN5 = PDFName(5)
 TxTBN6 = PDFName(6)
 TxTBN7 = PDFName(7)
 TxTBN8 = PDFName(8)
 TxTBN9 = PDFName(9)
TxTBN10 = PDFName(10)
TxTBN11 = PDFName(11)
TxTBN12 = PDFName(12)
TxTBN13 = PDFName(13)
TxTBN14 = PDFName(14)
TxTBN15 = PDFName(15)

Set objAttachments = Nothing
        Set objMsg = Nothing
  Set objSelection = Nothing
         Set objOL = Nothing
End Function
Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
         
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
     
    Set objApp = Nothing
End Function
Public Function AAPJEPUNTJE(BdNm, INTERN, SN, SubTxT, EM, t, WaardenEM, WaardenBdNm)

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
Waarden.EM.ControlTipText = "Mail van Collega"
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

Waarden.EM.ControlTipText = ""
Waarden.BdNm.ControlTipText = ""

AAPJEPUNTJE = BdNm

        Exceptions = strBdNM   '<------ finance processing firm
        If InStr(1, Exceptions, "|" & BdNm & "|", vbTextCompare) Then
        WaardenBdNm = True
        Waarden.BdNm.ControlTipText = "Extern financieel administratiekantoor"
        End If
        
        Exceptions = strEM   '<------ bounce e-mail address
        If InStr(1, Exceptions, "|" & EM & "|", vbTextCompare) Then
        WaardenEM = True
        Waarden.EM.ControlTipText = "Ongeldig e-mailadres"
        End If
        
        Exceptions = strEM1   '<------ noreply
        If InStr(1, Exceptions, "|" & EM1 & "|", vbTextCompare) Then
        WaardenEM = True
        Waarden.EM.ControlTipText = "no-reply e-mailadres"
        End If
        
Next


End Function
Sub ReplyWithAttachments()
    Dim rpl As Outlook.MailItem
    Dim itm As Object
     
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set rpl = itm.Reply
        CopyAttachments itm, rpl
        rpl.Display
    End If
     
    Set rpl = Nothing
    Set itm = Nothing
End Sub
 
Public Function GetAttachmentInfo(Attachment As Attachment)
    Dim Report
    
    If UCase(Right(Attachment.FileName, 3)) = "PDF" Then
    
    GetAttachmentInfo = ""
    
    Report = Report & Attachment.FileName
    
    GetAttachmentInfo = Report
    
    End If
    
    
End Function
Public Function GetAttachmentInfo2(Attachment As Attachment)
    Dim Report
    
    
    If UCase(Right(Attachment.FileName, 3)) = "PDF" Then
    
    GetAttachmentInfo2 = ""
    
    Report = Report & (Attachment.FileName)
        
    GetAttachmentInfo2 = Report
    End If
    
    
End Function
Sub CopyAttachments(objSourceItem, objTargetItem)
   Set FSO = CreateObject("Scripting.FileSystemObject")
   Set fldTemp = FSO.GetSpecialFolder(2) ' TemporaryFolder
   strPath = fldTemp & "\"
   
   For Each objAtt In objSourceItem.Attachments
   
   'If UCase(Right(objAtt.FileName, 3)) = "PDF" Then
   
   'If objAtt.Size > 4000 Then

      strFile = strPath & objAtt.FileName
      objAtt.SaveAsFile strFile
      
      objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
      
      FSO.DeleteFile strFile
    
    'End If
    
   Next
 
   Set fldTemp = Nothing
   Set FSO = Nothing
End Sub
Sub CopyAttachments2(objSourceItem, objTargetItem)
   Set FSO = CreateObject("Scripting.FileSystemObject")
   Set fldTemp = FSO.GetSpecialFolder(2) ' TemporaryFolder
   strPath = fldTemp & "\"
   For Each objAtt In objSourceItem.Attachments
            
   If objAtt.Size > 7200 Or UCase(Right(objAtt.FileName, 3)) = "PDF" Then

      strFile = strPath & objAtt.FileName
      objAtt.SaveAsFile strFile
      objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
      FSO.DeleteFile strFile
         
    End If
    
   Next
 
   Set fldTemp = Nothing
   Set FSO = Nothing
End Sub
Sub CopyAttachmentsFM(objSourceItem, objTargetItem)

   Set FSO = CreateObject("Scripting.FileSystemObject")
   Set fldTemp = FSO.GetSpecialFolder(2) ' TemporaryFolder
   strPath = fldTemp & "\"

   For Each objAtt In objSourceItem.Attachments

   'If objAtt.Size > 2001 Then
      strFile = strPath & objAtt.FileName
      objAtt.SaveAsFile strFile
      objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
      FSO.DeleteFile strFile
   'End If

   Next

   Set fldTemp = Nothing
   Set FSO = Nothing

End Sub
Sub SaveAtt2()
Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String

    strFolderpath = ("H:\Mijn Documenten\merge\pdf")
    On Error Resume Next

    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    strFolderpath = strFolderpath & "\OLAttachments\watermerk\"

    For Each objMsg In objSelection
    
    Set objAttachments = objMsg.Attachments
    lngCount = objAttachments.Count
        
    If lngCount > 0 Then
  
    For i = lngCount To 1 Step -1
    
        If UCase(Right(objAttachments.Item(i).FileName, 3)) = "PDF" Then
    
    strFile = objAttachments.Item(i).FileName
    strFile = strFolderpath & strFile
    objAttachments.Item(i).SaveAsFile strFile
    
        End If
    
    Next i
    
    End If
    
    Next
    
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
End Sub
Sub SaveAttCR()
Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String

    strFolderpath = ("G:\FIN\11DebCred\Crediteuren\20. Verwerking facturen\231. Creditfacturen")
    On Error Resume Next

    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    strFolderpath = strFolderpath & "\CREDIT NOTAS\"

    For Each objMsg In objSelection

    Set objAttachments = objMsg.Attachments
    lngCount = objAttachments.Count
        
    If lngCount > 0 Then
    
    For i = lngCount To 1 Step -1
    

    strFile = objAttachments.Item(i).FileName
    strFile = strFolderpath & strFile
    objAttachments.Item(i).SaveAsFile strFile
    
    Next i
    End If
    
    Next
    
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
End Sub
Sub VerwijderBijlages()
     
'Oude module waarmee manueel verstorende bijlage konden worden verwijderd, module is inactief
     
    Dim oExplorer As Outlook.Explorer
    Dim oMail As MailItem
    Set oExplorer = Application.ActiveExplorer
    Set oMail = oExplorer.Selection.Item(1).Forward
     
    On Error GoTo Release
    
    i = InputBox("Factuurnummer")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
     
    If oExplorer.Selection.Item(1).Class = olMail Then
        oMail.Subject = oMail.Subject & " " & i
        'oMail.HTMLBody = "Custom Text.<p> <img src=""custom image link""" _
        & " title=""D"" alt=""D"" name=""D"" border=""0"" id=""D""/>" _
        & vbCrLf & oMail.HTMLBody
        oMail.SentOnBehalfOfName = "facturen@amsterdam.nl"
        oMail.Recipients.Add "srvc47ACAM@amsterdam.nl"
        oMail.Recipients.Item(1).Resolve
        If oMail.Recipients.Item(1).Resolved Then
       
             oMail.DeferredDeliveryTime = DateAdd("s", 25, Now)
             oMail.Display
            'oMail.Save
            'oMail.Send
        Else
            MsgBox "Could not resolve " & oMail.Recipients.Item(1).Address
        End If
    Else
        MsgBox ("Not a mail item")
    End If
Release:
    Set oMail = Nothing
    Set oExplorer = Nothing
    
    Call KNOP1 '< aansturing voor registratie in Medewerkersformulier
    
    Call Afgehandeld
End Sub


