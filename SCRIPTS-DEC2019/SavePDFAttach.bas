Attribute VB_Name = "SavePDFAttach"
Sub SavePDF()

Dim objOL As Outlook.Application
Dim objItem As Outlook.MailItem
Dim objAttachments As Outlook.Attachments
Dim objAttachment As Outlook.Attachment
Dim objSelection As Outlook.Selection
Dim PDFCount As Long

strUser = Left(Environ("USERNAME"), 3)

Set objOL = New Outlook.Application
Dim olNs As Outlook.NameSpace
Dim oFolder As Outlook.MAPIFolder

    Set olNs = objOL.GetNamespace("MAPI")
    Set oFolder = Application.ActiveExplorer.CurrentFolder
    'Set oFolder = olNs.PickFolder

    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    
    
    Dim strFile As String
    
    Dim strFolderpath As String


 Dim currentExplorer As Explorer
 Dim Selection As Selection
 Dim olItem As Outlook.MailItem
 Dim itm As Object
 
Set currentExplorer = Application.ActiveExplorer
Set Selection = currentExplorer.Selection

strFolderpath = "H:\Mijn documenten\temp001"

'Result = MsgBox("Controleer of bovenste mail in map geselecteerd is" & vbLf & vbLf & "Mails in map: -" & oFolder & "- behandelen?", vbYesNo, "OnePDFSend")
'If Result = vbNo Then Exit Sub

    For Each objItem In oFolder.Items

    Set olItem = objItem
    

                NoPDFCount = 0
    '---------------------------------------------------begin code------------------------
    
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
Next
'End Funciton


'Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem
'Dim objAttachments As Outlook.Attachments
'Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
'Dim strFile As String
'Dim strFolderpath As String
'Dim strDeletedFiles As String

    strFolderpath = ("H:\Mijn documenten\temp001")
    On Error Resume Next

    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    strFolderpath = strFolderpath & "H:\Mijn documenten\temp001"

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
    
'ExitSub:
'
'Set objAttachments = Nothing
'Set objMsg = Nothing
'Set objSelection = Nothing
'Set objOL = Nothing
''End Sub

    
    'ICT blokkeert automailen: work-around *--------------------------------------
    '
    
    currentExplorer.Activate

    SendKeys "{Down}": DoEvents
    
    
    
    'End work-around *------------------------------------------------------------
        
   
    '-------------------------------------------------------^^^^^^ einde code ^^^^^^ ------------------------
    
    Next objItem
    
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
'End Sub

        
    Set fwd = Nothing
    Set itm = Nothing
        
        
End Sub
        



