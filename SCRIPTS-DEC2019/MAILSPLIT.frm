VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MAILSPLIT 
   Caption         =   "BESTANDEN IN MAIL PLAATSEN"
   ClientHeight    =   12675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   OleObjectBlob   =   "MAILSPLIT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MAILSPLIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Image4_Click()

Dim objOutlook, objnSpace, objCopyFolder As Object
Dim objItem, objCopy As Outlook.MailItem

Dim Attachments() As String
Dim i As Integer

    Dim CbBl(1 To 25) As String
    Dim CbFCT(1 To 25) As String
    Dim TxTBN(1 To 25) As String
    Dim FNTxT(1 To 25) As String

Set FSO = CreateObject("Scripting.FileSystemObject")
   
If MAILSPLIT.CbBl1 = True Or MAILSPLIT.CbFCT1 = True Then FilePathToAdd = strFolderpath & TxTBN1 & ","
If MAILSPLIT.CbFCT1 = True Then FN = FN & " & " & FNTxT1

If MAILSPLIT.CbBl2 = True Or MAILSPLIT.CbFCT2 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN2 & ","
If MAILSPLIT.CbFCT2 = True Then FN = FN & " & " & FNTxT2

If MAILSPLIT.CbBl3 = True Or MAILSPLIT.CbFCT3 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN3 & ","
If MAILSPLIT.CbFCT3 = True Then FN = FN & " & " & FNTxT3

If MAILSPLIT.CbBl4 = True Or MAILSPLIT.CbFCT4 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN4 & ","
If MAILSPLIT.CbFCT4 = True Then FN = FN & " & " & FNTxT4

If MAILSPLIT.CbBl5 = True Or MAILSPLIT.CbFCT5 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN5 & ","
If MAILSPLIT.CbFCT5 = True Then FN = FN & " & " & FNTxT5

If MAILSPLIT.CbBl6 = True Or MAILSPLIT.CbFCT6 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN6 & ","
If MAILSPLIT.CbFCT6 = True Then FN = FN & " & " & FNTxT6

If MAILSPLIT.CbBl7 = True Or MAILSPLIT.CbFCT7 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN7 & ","
If MAILSPLIT.CbFCT7 = True Then FN = FN & " & " & FNTxT7

If MAILSPLIT.CbBl8 = True Or MAILSPLIT.CbFCT8 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN8 & ","
If MAILSPLIT.CbFCT8 = True Then FN = FN & " & " & FNTxT8

If MAILSPLIT.CbBl9 = True Or MAILSPLIT.CbFCT9 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN9 & ","
If MAILSPLIT.CbFCT9 = True Then FN = FN & " & " & FNTxT9

If MAILSPLIT.CbBl10 = True Or MAILSPLIT.CbFCT10 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN10 & ","
If MAILSPLIT.CbFCT10 = True Then FN = FN & " & " & FNTxT10

If MAILSPLIT.CbBl11 = True Or MAILSPLIT.CbFCT11 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN11 & ","
If MAILSPLIT.CbFCT11 = True Then FN = FN & " & " & FNTxT11

If MAILSPLIT.CbBl12 = True Or MAILSPLIT.CbFCT12 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN12 & ","
If MAILSPLIT.CbFCT12 = True Then FN = FN & " & " & FNTxT12

If MAILSPLIT.CbBl13 = True Or MAILSPLIT.CbFCT13 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN13 & ","
If MAILSPLIT.CbFCT13 = True Then FN = FN & " & " & FNTxT13

If MAILSPLIT.CbBl14 = True Or MAILSPLIT.CbFCT14 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN14 & ","
If MAILSPLIT.CbFCT14 = True Then FN = FN & " & " & FNTxT14

If MAILSPLIT.CbBl15 = True Or MAILSPLIT.CbFCT15 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN15 & ","
If MAILSPLIT.CbFCT15 = True Then FN = FN & " & " & FNTxT15

If MAILSPLIT.CbBl16 = True Or MAILSPLIT.CbFCT16 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN16 & ","
If MAILSPLIT.CbFCT16 = True Then FN = FN & " & " & FNTxT16

If MAILSPLIT.CbBl17 = True Or MAILSPLIT.CbFCT17 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN17 & ","
If MAILSPLIT.CbFCT17 = True Then FN = FN & " & " & FNTxT17

If MAILSPLIT.CbBl18 = True Or MAILSPLIT.CbFCT18 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN18 & ","
If MAILSPLIT.CbFCT18 = True Then FN = FN & " & " & FNTxT18

If MAILSPLIT.CbBl19 = True Or MAILSPLIT.CbFCT19 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN19 & ","
If MAILSPLIT.CbFCT19 = True Then FN = FN & " & " & FNTxT19

If MAILSPLIT.CbBl20 = True Or MAILSPLIT.CbFCT20 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN20 & ","
If MAILSPLIT.CbFCT20 = True Then FN = FN & " & " & FNTxT20

If MAILSPLIT.CbBl21 = True Or MAILSPLIT.CbFCT21 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN21 & ","
If MAILSPLIT.CbFCT21 = True Then FN = FN & " & " & FNTxT21

If MAILSPLIT.CbBl22 = True Or MAILSPLIT.CbFCT22 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN22 & ","
If MAILSPLIT.CbFCT22 = True Then FN = FN & " & " & FNTxT22

If MAILSPLIT.CbBl23 = True Or MAILSPLIT.CbFCT23 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN23 & ","
If MAILSPLIT.CbFCT23 = True Then FN = FN & " & " & FNTxT23

If MAILSPLIT.CbBl24 = True Or MAILSPLIT.CbFCT24 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN24 & ","
If MAILSPLIT.CbFCT24 = True Then FN = FN & " & " & FNTxT24

If MAILSPLIT.CbBl25 = True Or MAILSPLIT.CbFCT25 = True Then FilePathToAdd = FilePathToAdd & strFolderpath & TxTBN25 & ","
If MAILSPLIT.CbFCT25 = True Then FN = FN & " & " & FNTxT25

FilePathToAdd = Left(FilePathToAdd, Len(FilePathToAdd) - 1)

FN = Right(FN, Len(FN) - 3)
   
Set objOutlook = CreateObject("Outlook.Application")
Set objnSpace = objOutlook.GetNamespace("MAPI")
Set objCopyFolder = Application.ActiveExplorer.CurrentFolder '.Folders("Copy")
Set objItem = Application.ActiveExplorer.Selection.Item(1)

Set objCopy = objItem.Copy

Set objCopy = Application.ActiveExplorer.Selection.Item(1)

        Do Until objCopy.Attachments.Count = 0
            objCopy.Attachments.Remove (1)
        Loop
        
objCopy.Subject = objCopy.Subject & " (c) " & FN
objCopy.UnRead = True
            
        If FilePathToAdd <> "" Then
            Attachments = Split(FilePathToAdd, ",")
            For i = LBound(Attachments) To UBound(Attachments)
                If Attachments(i) <> "" Then
            
objCopy.Attachments.Add Trim(Attachments(i))
            
                End If
            Next i
         End If
objCopy.Save

MAILSPLIT.Hide

KillAll

   Set fldTemp = Nothing
   Set FSO = Nothing
Set objFolder = Nothing
Set objnSpace = Nothing
Set objOutlook = Nothing
Set objItem = Nothing
Set objCopy = Nothing

End Sub
Private Sub Close1_Click()
Me.Caption = "PDF bestanden uit map verwijderen....."
KillAll
Me.Caption = "BESTANDEN IN NIEUWE MAIL PLAATSEN"
MAILSPLIT.Hide
End Sub
Private Sub Image5_Click()

Dim objOutlook, objnSpace, objCopyFolder As Object
Dim objItem, objCopy As Outlook.MailItem

Dim Attachment() As String
Dim i As Integer

Path = ("H:\Mijn Documenten\merge\pdf")
Path = Path & "\OLAttachments\"

FileExt = "*.pdf*"

FilePathToAdd = Dir(Path & FileExt)

Do While Len(FilePathToAdd) > 0
   
Set objOutlook = CreateObject("Outlook.Application")
Set objnSpace = objOutlook.GetNamespace("MAPI")
Set objCopyFolder = Application.ActiveExplorer.CurrentFolder '.Folders("Copy")
Set objItem = Application.ActiveExplorer.Selection.Item(1)

Subject1 = objItem.Subject

Set objCopy = objItem.Copy

Set objCopy = Application.ActiveExplorer.Selection.Item(1)

        Do Until objCopy.Attachments.Count = 0
            objCopy.Attachments.Remove (1)
        Loop
        
objCopy.Subject = FilePathToAdd & " (c)"
objCopy.UnRead = True
            
        If FilePathToAdd <> "" Then
            Attachment = Split(Path & FilePathToAdd, ",")
            For i = LBound(Attachment) To UBound(Attachment)
                If Attachment(i) <> "" Then
            
objCopy.Attachments.Add Trim(Attachment(i))
            
                End If
            Next i
         End If
objCopy.Save

Kill (Path & FilePathToAdd)

FilePathToAdd = Dir()

 Loop

MAILSPLIT.Hide

   Set fldTemp = Nothing
   Set FSO = Nothing
Set objFolder = Nothing
Set objnSpace = Nothing
Set objOutlook = Nothing
Set objItem = Nothing
Set objCopy = Nothing

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 1 Then Cancel = 1
End Sub
