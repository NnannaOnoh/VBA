Attribute VB_Name = "EMAILTELLING"
Sub tellingPOSTVAKIN()

MBOX = "Facturen"
FOLDERNM = "Postvak IN"
CsvNM = "POSTVAKIN.csv"

Call telling(MBOX, FOLDERNM, CsvNM)

End Sub
Sub tellingNieuwFacturen()

MBOX = "Facturen"
FOLDERNM = ">Nieuwe Facturen"
CsvNM = "NIEUWEFACTUREN.csv"

Call SFtelling(MBOX, FOLDERNM, CsvNM)

End Sub
Sub tellingINCASSO()

MBOX = "Facturen"
FOLDERNM = "0000Automatische incasso zonder routecode"
CsvNM = "INCASSO.csv"

Call SFtelling(MBOX, FOLDERNM, CsvNM)

End Sub
Sub tellingCREDIT()

MBOX = "Facturen"
FOLDERNM = "0000Creditnota uitzoeken"
CsvNM = "CREDIT.csv"

Call SFtelling(MBOX, FOLDERNM, CsvNM)

End Sub
Sub tellingAANMANING()

MBOX = "Facturen"
FOLDERNM = "00Aanmaningen"
CsvNM = "AANMANINGEN.csv"

Call SFtelling(MBOX, FOLDERNM, CsvNM)

End Sub
Sub tellingHERINNERING()

MBOX = "Facturen"
FOLDERNM = "00Herinneringen"
CsvNM = "HERINNERING.csv"

Call SFtelling(MBOX, FOLDERNM, CsvNM)

End Sub
Sub telling(MBOX, FOLDERNM, CsvNM)

    Dim objOutlook As Object, objnSpace As Object, objFolder As MAPIFolder
    Dim EmailCount As Integer
    Set objOutlook = CreateObject("Outlook.Application")
    Set objnSpace = objOutlook.GetNamespace("MAPI")
      
        On Error Resume Next
        Set objFolder = objnSpace.folders(MBOX).folders(FOLDERNM)
        If Err.Number <> 0 Then
        Err.Clear
        MsgBox "No such folder."
        Exit Sub
        End If

    EmailCount = objFolder.Items.Count

    MsgBox "Number of emails in the folder: " & EmailCount, , "email count"

    Dim dateStr As String
    Dim myItems As Outlook.Items
    Dim dict As Object
    Dim msg As String
    Set dict = CreateObject("Scripting.Dictionary")
    Set myItems = objFolder.Items
    myItems.SetColumns ("SentOn")
    ' Determine date of each message:
    For Each MyItem In myItems
        dateStr = GetDate1(MyItem.SentOn)
        If Not dict.Exists(dateStr) Then
            dict (dateStr) >= 0
        End If
        dict(dateStr) = CLng(dict(dateStr)) + 1
    Next MyItem
    
    ' Output counts per day:
    msg = ""
    For Each o In dict.Keys
        msg = msg & o & "; " & dict(o) & vbCrLf
    Next
    MsgBox msg
        
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
Set oFile = FSO.CreateTextFile("G:\FIN\11DebCred\Crediteuren\60. Team Input\KPI-formulieren Team Input (voor op de werkvloer)\Digitale formulieren\Vandaag\OUTLOOK LOG\" & CsvNM)

oFile.WriteLine msg
oFile.Close
Set FSO = Nothing
Set oFile = Nothing
    
   
    Set objFolder = Nothing
    Set objnSpace = Nothing
    Set objOutlook = Nothing
    
    End Sub
    Sub SFtelling(MBOX, FOLDERNM, CsvNM)

    Dim objOutlook As Object, objnSpace As Object, objFolder As MAPIFolder
    Dim EmailCount As Integer
    Set objOutlook = CreateObject("Outlook.Application")
    Set objnSpace = objOutlook.GetNamespace("MAPI")
      
        On Error Resume Next
        Set objFolder = objnSpace.folders(MBOX).folders("Postvak IN").folders(FOLDERNM)
        If Err.Number <> 0 Then
        Err.Clear
        MsgBox "No such folder."
        Exit Sub
        End If

    EmailCount = objFolder.Items.Count

    MsgBox "Number of emails in the folder: " & EmailCount, , "email count"

    Dim dateStr As String
    Dim myItems As Outlook.Items
    Dim dict As Object
    Dim msg As String
    Set dict = CreateObject("Scripting.Dictionary")
    Set myItems = objFolder.Items
    myItems.SetColumns ("SentOn")
    ' Determine date of each message:
    For Each MyItem In myItems
        dateStr = GetDate1(MyItem.SentOn)
        If Not dict.Exists(dateStr) Then
            dict (dateStr) >= 0
        End If
        dict(dateStr) = CLng(dict(dateStr)) + 1
    Next MyItem
    
    ' Output counts per day:
    msg = ""
    For Each o In dict.Keys
        msg = msg & o & "; " & dict(o) & vbCrLf
    Next
    MsgBox msg
        
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
Set oFile = FSO.CreateTextFile("G:\FIN\11DebCred\Crediteuren\60. Team Input\KPI-formulieren Team Input (voor op de werkvloer)\Digitale formulieren\Vandaag\OUTLOOK LOG\" & CsvNM)

oFile.WriteLine msg
oFile.Close
Set FSO = Nothing
Set oFile = Nothing
    
   
    Set objFolder = Nothing
    Set objnSpace = Nothing
    Set objOutlook = Nothing
    
    End Sub
    Function GetDate1(dt As Date) As String
    GetDate1 = Year(dt) & "-" & Month(dt) & "-" & Day(dt)
End Function
