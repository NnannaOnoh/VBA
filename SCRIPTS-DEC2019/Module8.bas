Attribute VB_Name = "Module8"
Sub CopyToExcel1()
 Dim xlApp As Object
 Dim xlWB As Object
 Dim xlSheet As Object
 Dim rCount As Long
 Dim bXStarted As Boolean
 Dim enviro As String
 Dim strPath As String
 
Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem
Dim objAttachments As Outlook.Attachments
Dim objAttachment As Outlook.Attachment
Dim objSelection As Outlook.Selection
Dim i, a As Long
Dim lngCount, PDFCount As Long


Set objOL = New Outlook.Application
Dim olNs As Outlook.NameSpace
Dim oFolder As Outlook.MAPIFolder

    Set olNs = objOL.GetNamespace("MAPI")
    Set oFolder = olNs.PickFolder

    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection


 Dim currentExplorer As Explorer
 Dim Selection As Selection
 Dim olItem As Outlook.MailItem
 Dim obj As Object
 Dim strColA, strColB, strColC, strColD, strColE As String
               
     On Error Resume Next
     Set xlApp = GetObject(, "Excel.Application")
     If Err <> 0 Then
         objOL.StatusBar = "Please wait while Excel source is opened ... "
         Set xlApp = CreateObject("Excel.Application")
         bXStarted = True
     End If
     On Error GoTo 0

Set xlWB = xlApp.Workbooks.Open("H:\Mijn Documenten\merge\pdf\outlook export.xlsm")
Set xlSheet = xlWB.Sheets("Outlook")

  xlSheet.Range("A1") = "Recieved Time"
  xlSheet.Range("B1") = "Sender"
  xlSheet.Range("C1") = "Sender address"
  xlSheet.Range("D1") = "Subject"
  xlSheet.Range("E1") = "Sent To"
  xlSheet.Range("F1") = "PDF Attachments"
  xlSheet.Range("G1") = "XML Attachments"
  xlSheet.Range("H1") = "XLS Attachments"
  xlSheet.Range("I1") = "CSV Attachments"
  xlSheet.Range("J1") = "HTM Attachments"
  xlSheet.Range("K1") = "DOC Attachments"
  xlSheet.Range("L1") = "JPG Attachments"
  xlSheet.Range("M1") = "Total Attachments"
  xlSheet.Range("N1") = "Size"
  On Error Resume Next
  
rCount = xlSheet.Range("A" & xlSheet.Rows.Count).End(-4162).Row
rCount = rCount + 1

Set currentExplorer = Application.ActiveExplorer
Set Selection = currentExplorer.Selection

    Set olItem = objItem
    Set objAttachments = objItem.Attachments
        

        
        For Each objItem In oFolder.Items
        
                PDFCount = 0
                XMLCount = 0
                XLSCount = 0
                CSVCount = 0
                HTMCount = 0
                DOCCount = 0
                JPGCount = 0
                AttachmentCount = 0
                
            If objItem.Attachments.Count <> 0 Then
                For Each objAttachment In objItem.Attachments
                    If UCase(Right(objAttachment.FileName, 3)) = "PDF" Then
                        a = a + 1
                        PDFCount = PDFCount + 1
                 End If
                 Next
                For Each objAttachment In objItem.Attachments
                    If UCase(Right(objAttachment.FileName, 3)) = "XML" Then
                        a = a + 1
                        XMLCount = XMLCount + 1
                    End If
                 Next
                For Each objAttachment In objItem.Attachments
                    If UCase(Right(objAttachment.FileName, 3)) = "XLS" Or UCase(Right(objAttachment.FileName, 4)) = "XLSX" Then
                        a = a + 1
                        XLSCount = XLSCount + 1
                    End If
                 Next
                For Each objAttachment In objItem.Attachments
                    If UCase(Right(objAttachment.FileName, 3)) = "CSV" Then
                        a = a + 1
                        CSVCount = CSVCount + 1
                    End If
                 Next
                For Each objAttachment In objItem.Attachments
                    If UCase(Right(objAttachment.FileName, 3)) = "HTM" Then
                        a = a + 1
                        HTMCount = HTMCount + 1
                    End If
                 Next
                For Each objAttachment In objItem.Attachments
                    If UCase(Right(objAttachment.FileName, 3)) = "DOC" Or UCase(Right(objAttachment.FileName, 4)) = "DOCX" Then
                        a = a + 1
                        DOCCount = DOCCount + 1
                    End If
                 Next
                For Each objAttachment In objItem.Attachments
                    If UCase(Right(objAttachment.FileName, 3)) = "JPG" Then
                        a = a + 1
                        JPGCount = JPGCount + 1
                    End If
                 Next
                For Each objAttachment In objItem.Attachments
                        a = a + 1
                        AttachmentCount = AttachmentCount + 1
                    
                    Next objAttachment
            End If
        
Application.ScreenUpdating = False
    
    If objItem.Class = olMail Then

    strColA = objItem.ReceivedTime
    strColB = objItem.SenderName
    strColC = objItem.SenderEmailAddress
    strColD = objItem.Subject
    strColE = objItem.To
    strColF = PDFCount
    strColG = XMLCount
    strColH = XLSCount
    strColI = CSVCount
    strColJ = HTMCount
    strColK = DOCCount
    strColL = JPGCount
    strColM = AttachmentCount
    strColN = objItem.Size

Dim strRecipients As String
Dim Recipient As Outlook.Recipient
For Each Recipient In olItem.Recipients
 strRecipients = Recipient.Address & "; " & strRecipients
 Next Recipient

  strColG = objItem.strRecipients

 Dim olEU As Outlook.ExchangeUser
 Dim oEDL As Outlook.ExchangeDistributionList
 Dim recip As Outlook.Recipient
 Set recip = Application.Session.CreateRecipient(strColB)

If InStr(1, strColB, "/") > 0 Then

    Select Case recip.AddressEntry.AddressEntryUserType
       Case OlAddressEntryUserType.olExchangeUserAddressEntry
         Set olEU = recip.AddressEntry.GetExchangeUser
         If Not (olEU Is Nothing) Then
             strColB = olEU.PrimarySmtpAddress
         End If
       Case OlAddressEntryUserType.olOutlookContactAddressEntry
         Set olEU = recip.AddressEntry.GetExchangeUser
         If Not (olEU Is Nothing) Then
            strColB = olEU.PrimarySmtpAddress
         End If
       Case OlAddressEntryUserType.olExchangeDistributionListAddressEntry
         Set oEDL = recip.AddressEntry.GetExchangeDistributionList
         If Not (oEDL Is Nothing) Then
            strColB = olEU.PrimarySmtpAddress
         End If
     End Select
End If

  xlSheet.Range("A" & rCount) = strColA
  xlSheet.Range("B" & rCount) = strColB
  xlSheet.Range("C" & rCount) = strColC
  xlSheet.Range("D" & rCount) = strColD
  xlSheet.Range("E" & rCount) = strColE
  xlSheet.Range("F" & rCount) = strColF
  xlSheet.Range("G" & rCount) = strColG
  xlSheet.Range("H" & rCount) = strColH
  xlSheet.Range("I" & rCount) = strColI
  xlSheet.Range("J" & rCount) = strColJ
  xlSheet.Range("K" & rCount) = strColK
  xlSheet.Range("L" & rCount) = strColL
  xlSheet.Range("M" & rCount) = strColM
  xlSheet.Range("N" & rCount) = strColN
  
  rCount = rCount + 1


strRecipients = ""

    End If
    
Application.ScreenUpdating = True

    Next
 
 xlApp.Visible = True

xlApp.Run "RunCalc"

     Set olItem = Nothing
     Set obj = Nothing
     Set currentExplorer = Nothing
     Set xlSheet = Nothing
     Set xlWB = Nothing
     Set xlApp = Nothing
     
     MsgBox "Done " & rCount - 2 & " mails."
 End Sub
