Attribute VB_Name = "ATTAC"
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
 
Public Function GetAttachmentInfo(attachment As attachment)
    Dim Report
    
    
    If Right(attachment.FileName, 3) = "pdf" Or Right(attachment.FileName, 3) = "PDF" Then
    
    GetAttachmentInfo = ""
    
    Report = Report & Left(attachment.FileName, Len(attachment.FileName) - 4)
        
    GetAttachmentInfo = Report
    End If
    
    
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
 
Sub CopyAttachments(objSourceItem, objTargetItem)
   Set FSO = CreateObject("Scripting.FileSystemObject")
   Set fldTemp = FSO.GetSpecialFolder(2) ' TemporaryFolder
   strPath = fldTemp & "\"
   
   For Each objAtt In objSourceItem.Attachments
   
   If Right(objAtt.FileName, 3) = "pdf" Or Right(objAtt.FileName, 3) = "PDF" Then

      strFile = strPath & objAtt.FileName
      objAtt.SaveAsFile strFile
      objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
      FSO.DeleteFile strFile
    
    End If
    
   Next
 
   Set fldTemp = Nothing
   Set FSO = Nothing
End Sub
 
Sub CopyAttachments2(objSourceItem, objTargetItem)
   Set FSO = CreateObject("Scripting.FileSystemObject")
   Set fldTemp = FSO.GetSpecialFolder(2) ' TemporaryFolder
   strPath = fldTemp & "\"
   For Each objAtt In objSourceItem.Attachments
            
   If objAtt.Size > 7200 Or Right(objAtt.FileName, 3) = "PDF" Or "pdf" Then

      strFile = strPath & objAtt.FileName
      objAtt.SaveAsFile strFile
      objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
      FSO.DeleteFile strFile
    
         
    End If
    
   Next
 
   Set fldTemp = Nothing
   Set FSO = Nothing
End Sub
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
    
    Call Afgehandeld
    
End Sub


