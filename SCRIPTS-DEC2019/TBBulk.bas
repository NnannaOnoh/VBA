Attribute VB_Name = "TBBulk"
Option Explicit
 Sub TBB()
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
 Dim strColB, strColC, strColD, strColE, strColA, strColF, Path, FileName, FolderName As String
 Dim Controle As Integer
 Dim Controle1 As Integer

FolderName = "G:\FIN\11DebCred\Crediteuren\20. Verwerking facturen\260. Teruggestuurde facturen door Input"

Shell "C:\WINDOWS\explorer.exe """ & FolderName & "", vbNormalFocus



    Controle = MsgBox("Controleer of bestand in gebruik is. Doorgaan?", vbYesNo + vbQuestion, "Controleer bestand")

    If Controle = vbNo Then Exit Sub
    
    Controle1 = MsgBox("Is Excel gestart?", vbYesNo + vbQuestion, "Start Excel")

    If Controle1 = vbNo Then Exit Sub
    
    Path = Path = "G:\FIN\11DebCred\Crediteuren\20. Verwerking facturen\260. Teruggestuurde facturen door Input\"
    FileName = Format(Now, "yyyy-mm-dd")

 strPath = "G:\FIN\11DebCred\Crediteuren\20. Verwerking facturen\260. Teruggestuurde facturen door Input\TBB\Teruggestuurde Facturen BULK2.xlsx"
     On Error Resume Next
     Set xlApp = GetObject(, "Excel.Application")
     If Err <> 0 Then
         Application.StatusBar = "Please wait while Excel source is opened ... "
         Set xlApp = CreateObject("Excel.Application")
         bXStarted = True
     End If
     On Error GoTo 0
     
xlApp.Application.ScreenUpdating = True

     'Open the workbook to input the data
     Set xlWB = xlApp.Workbooks.Open(strPath)
     Set xlSheet = xlWB.Sheets(1)
    ' Process the message record
    
    On Error Resume Next
'Find the next empty line of the worksheet
rCount = xlSheet.Range("B" & xlSheet.Rows.Count).End(-4162).Row
'needed for Exchange 2016. Remove if causing blank lines.
rCount = rCount + 1

' get the values from outlook
Set currentExplorer = Application.ActiveExplorer
Set Selection = currentExplorer.Selection
  For Each obj In Selection

    Set olItem = obj
    
 'collect the fields
    strColB = olItem.SenderName
    strColC = olItem.SenderEmailAddress
    'strColD = olItem.Body
    strColD = olItem.To
    strColA = olItem.ReceivedTime
    strColE = olItem.Subject

' Get the Exchange address
' if not using Exchange, this block can be removed
 Dim olEU As Outlook.ExchangeUser
 Dim oEDL As Outlook.ExchangeDistributionList
 Dim recip As Outlook.Recipient
 Set recip = Application.Session.CreateRecipient(strColC)

 If InStr(1, strColC, "/") > 0 Then
' if exchange, get smtp address
     Select Case recip.AddressEntry.AddressEntryUserType
       Case OlAddressEntryUserType.olExchangeUserAddressEntry
         Set olEU = recip.AddressEntry.GetExchangeUser
         If Not (olEU Is Nothing) Then
             strColC = olEU.PrimarySmtpAddress
         End If
       Case OlAddressEntryUserType.olOutlookContactAddressEntry
         Set olEU = recip.AddressEntry.GetExchangeUser
         If Not (olEU Is Nothing) Then
            strColC = olEU.PrimarySmtpAddress
         End If
       Case OlAddressEntryUserType.olExchangeDistributionListAddressEntry
         Set oEDL = recip.AddressEntry.GetExchangeDistributionList
         If Not (oEDL Is Nothing) Then
            strColC = olEU.PrimarySmtpAddress
         End If
     End Select
End If
' End Exchange section

'write them in the excel sheet
  xlSheet.Range("B" & rCount) = strColB
  xlSheet.Range("c" & rCount) = strColC
  'xlSheet.Range("d" & rCount) = strColD
  xlSheet.Range("d" & rCount) = strColD
  xlSheet.Range("a" & rCount) = strColA
  xlSheet.Range("e" & rCount) = strColE
 
'Next row
  rCount = rCount + 1

 Next

    '' xlWb.Close 1
     If bXStarted Then
         xlApp.Quit
     End If
     
xlApp.Application.ScreenUpdating = True
    
     Set olItem = Nothing
     Set obj = Nothing
     Set currentExplorer = Nothing
     Set xlApp = Nothing
     Set xlWB = Nothing
     Set xlSheet = Nothing
     
MsgBox "Uitgevoerd"
Act1:
 End Sub


