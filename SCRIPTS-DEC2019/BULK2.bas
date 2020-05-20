Attribute VB_Name = "BULK2"
 Sub CopyToExcel()
 
' Deze module is niet in gebruik
 
 Dim xlApp As Object
 Dim xlWB As Object
 Dim xlSheet As Object
 Dim rCount As Long
 Dim bXStarted As Boolean
 Dim enviro As String
 Dim strPath As String

 Dim currentExplorer As Explorer
 Dim Selection As Selection
 Dim olItem As Outlook.MailItem
 Dim obj As Object
 Dim strColA, strColB, strColC, strColD, strColE As String
 
     On Error Resume Next
     Set xlApp = GetObject(, "Excel.Application")
     If Err <> 0 Then
         Application.StatusBar = "Please wait while Excel source is opened ... "
         Set xlApp = CreateObject("Excel.Application")
         bXStarted = True
     End If
     On Error GoTo 0
     

Set xlWB = xlApp.Workbooks.Add
Set xlSheet = xlWB.Sheets("Blad1")

  xlSheet.Range("A1") = "Recieved Time"
  xlSheet.Range("B1") = "Sender"
  xlSheet.Range("C1") = "Sender address"
  xlSheet.Range("D1") = "Subject"
  xlSheet.Range("E1") = "Sent To"
    
  On Error Resume Next

rCount = xlSheet.Range("A" & xlSheet.Rows.Count).End(-4162).Row
rCount = rCount + 1

Set currentExplorer = Application.ActiveExplorer
Set Selection = currentExplorer.Selection
  For Each obj In Selection

    Set olItem = obj
    
    strColA = olItem.ReceivedTime
    strColB = olItem.SenderName
    strColC = olItem.SenderEmailAddress
    strColD = olItem.Subject
    strColE = olItem.To 'display name

Dim strRecipients As String
Dim Recipient As Outlook.Recipient
For Each Recipient In olItem.Recipients
 strRecipients = Recipient.Address & "; " & strRecipients
 Next Recipient

  strColE = strRecipients 'email address

 Dim olEU As Outlook.ExchangeUser
 Dim oEDL As Outlook.ExchangeDistributionList
 Dim recip As Outlook.Recipient
 Set recip = Application.Session.CreateRecipient(strColB)

If InStr(1, strColB, "/") > 0 Then
' if exchange, get smtp address
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

  xlSheet.Range("A" & rCount) = strColA ' sender name
  xlSheet.Range("B" & rCount) = strColB ' sender address
  xlSheet.Range("C" & rCount) = strColC ' subject
  xlSheet.Range("D" & rCount) = strColD ' sent to
  xlSheet.Range("E" & rCount) = strColE ' recieved time
  rCount = rCount + 1

strRecipients = ""

 Next
 xlApp.Visible = True

    
     Set olItem = Nothing
     Set obj = Nothing
     Set currentExplorer = Nothing
     Set xlSheet = Nothing
     Set xlWB = Nothing
     Set xlApp = Nothing
 End Sub
