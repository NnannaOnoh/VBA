Attribute VB_Name = "XML"
Sub GETXML1()

Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String
Dim EM As String

    strFolderpath = ("H:\Mijn Documenten\merge\pdf\Splitsen\")
   
    'On Error Resume Next

    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    'strFolderpath = strFolderpath & "\XML\februari\"

    For Each objMsg In objSelection
    
    EM = objMsg.SenderEmailAddress
    
    'EM = Left(EM, Len(EM) - 3)
    
    'MsgBox EM

    Set objAttachments = objMsg.Attachments
   
    lngCount = objAttachments.Count
            
        If lngCount > 0 Then
    
    For i = lngCount To 1 Step -1
    
    strFile = objAttachments.Item(i).FileName
                   
    'strFile = strFolderpath & EM & ".xml"
    
    'strFile = EM
    
    'strFile = Replace(strFile, ",", " ")
    
    'strFile = strFile & ".xml"
    
    'If Right(objAttachments.Item(i).FileName, 3) = "pdf" Or Right(objAttachments.Item(i).FileName, 3) = "PDF" Then objAttachments.Item(i).SaveAsFile strFolderpath & strFile
       
       
     If Right(objAttachments.Item(i).FileName, 3) = ("pdf") Then objAttachments.Item(i).SaveAsFile strFolderpath & strFile
     'Or Right(objAttachments.Item(i).FileName, 3) = "PDF"
      
    Next i
       
        End If
           
    Next


    
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing

MsgBox "DONE!"

End Sub
Sub GETXML()

Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long

Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String

Dim EM As String

Set xlApp = Excel.Application
Set objOL = New Outlook.Application
Dim olNs As Outlook.NameSpace
Dim oFolder As Outlook.MAPIFolder

Set olNs = objOL.GetNamespace("MAPI")

InitFolder = "G:\FIN\11DebCred\Crediteuren\24. Verwerking E-facturatie\XML"
   
    On Error Resume Next
    
    Set oFolder = olNs.PickFolder
    
    With xlApp.FileDialog(msoFileDialogFolderPicker)
        .Show
        InitFolder = .InitialFileName
        strFolderpath = .SelectedItems(1)
    End With
    
    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    
    For Each objItem In oFolder.Items
    
    If objItem.Class = olMail Then
    
    EM = objItem.SenderEmailAddress

    Set objAttachments = objItem.Attachments
   
    lngCount = objAttachments.Count
            
        If lngCount > 0 Then
    
    For i = lngCount To 1 Step -1
    
    strFile = objAttachments.Item(i).FileName
                   
    strFile = strFolderpath & "\" & EM & ".xml"
    
    If Right(objAttachments.Item(i).FileName, 3) = "xml" Or Right(objAttachments.Item(i).FileName, 3) = "XML" Then objAttachments.Item(i).SaveAsFile strFile
       
    Next i
       
        End If
        
    End If
           
    Next
    
    Set oFolder = oFolder.Folders("Retour leverancier")

    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    
    For Each objItem In oFolder.Items
    
    If objItem.Class = olMail Then
    
    EM = objItem.SenderEmailAddress

    Set objAttachments = objItem.Attachments
   
    lngCount = objAttachments.Count
            
        If lngCount > 0 Then
    
    For i = lngCount To 1 Step -1
    
    strFile = objAttachments.Item(i).FileName
                   
    strFile = strFolderpath & "\" & EM & ".xml"
    
    If Right(objAttachments.Item(i).FileName, 3) = "xml" Or Right(objAttachments.Item(i).FileName, 3) = "XML" Then objAttachments.Item(i).SaveAsFile strFile
       
    Next i
       
        End If
        
    End If
           
    Next

    
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing

MsgBox "DONE!"

End Sub
Sub GETPDF()

Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long

Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String

Dim EM As String

Set objOL = New Outlook.Application
Dim olNs As Outlook.NameSpace
Dim oFolder As Outlook.MAPIFolder

Set olNs = objOL.GetNamespace("MAPI")

    strFolderpath = ("H:\Mijn Documenten\merge\pdf\")
   
    On Error Resume Next
       
    Set oFolder = olNs.Folders("Facturen").Folders("Postvak IN").Folders("Nnanna").Folders("facturen voor kofax") '.Folders("Retour leverancier")

    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    strFolderpath = strFolderpath & "Splitsen" & "\"

    'For Each objMsg In objSelection
    
    For Each objItem In oFolder.Items
    
    If objItem.Class = olMail Then
    
    'EM = objItem.SenderEmailAddress
    
    'EM = Left(EM, Len(EM) - 3)
    
    'MsgBox EM

    Set objAttachments = objItem.Attachments
   
    lngCount = objAttachments.Count
            
        If lngCount > 0 Then
    
    For i = lngCount To 1 Step -1
    
    strFile = objAttachments.Item(i).FileName
                   
    'strFile = strFolderpath & EM & ".xml"
    
    'strFile = EM
    
    'strFile = Replace(strFile, ",", " ")
    
    'strFile = strFile & ".xml"
    
    If Right(objAttachments.Item(i).FileName, 3) = "pdf" Or Right(objAttachments.Item(i).FileName, 3) = "PDF" Then objAttachments.Item(i).SaveAsFile strFile
       
    Next i
       
        End If
        
    End If
           
    Next

    
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing

MsgBox "DONE!"

End Sub
