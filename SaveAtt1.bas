Attribute VB_Name = "SaveAtt1"
Sub SaveAtt()
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
    strFolderpath = strFolderpath & "\OLAttachments\"

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
