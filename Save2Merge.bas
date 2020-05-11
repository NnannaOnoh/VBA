Attribute VB_Name = "Save2Merge"
Sub SaveToMerge()
     
    Dim DestFile As String
    
    FN = InputBox("Factuurnummer")
    Select Case StrPtr(FN)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
    DestFile = FN & "(M)" & ".pdf" '
    
     
    Dim MyPath As String, MyFiles As String, ToPath As String
    
    Dim a() As String, i As Long, f As String
     

    MyPath = "H:\Mijn Documenten\merge\pdf\OLAttachments\"


    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    ReDim a(1 To 2 ^ 14)
    f = Dir(MyPath & "*.pdf")
    While Len(f)
        If StrComp(f, DestFile, vbTextCompare) Then
            i = i + 1
            a(i) = f
        End If
        f = Dir()
    Wend
     

    If i Then
        ReDim Preserve a(1 To i)
        MyFiles = Join(a, ",")

        Call MergePDFs(MyPath, MyFiles, DestFile)

    Else
        MsgBox "No PDF files found in" & vbLf & MyPath, vbExclamation, "Canceled"
    End If
     
End Sub
 
Sub MergePDFs(MyPath As String, MyFiles As String, Optional DestFile As String = "MergedFile.pdf")

     
    Dim a As Variant, i As Long, n As Long, ni As Long, p As String
    Dim AcroApp As New Acrobat.AcroApp, PartDocs() As Acrobat.CAcroPDDoc
     
    If Right(MyPath, 1) = "\" Then p = MyPath Else p = MyPath & "\"
    a = Split(MyFiles, ",")
    ReDim PartDocs(0 To UBound(a))
     
    On Error GoTo exit_
    If Len(Dir(p & DestFile)) Then Kill p & DestFile
    For i = 0 To UBound(a)

        If Dir(p & Trim(a(i))) = "" Then
            MsgBox "File not found" & vbLf & p & a(i), vbExclamation, "Canceled"
            Exit For
        End If

        Set PartDocs(i) = CreateObject("AcroExch.PDDoc")
        PartDocs(i).Open p & Trim(a(i))
        If i Then

            ni = PartDocs(i).GetNumPages()
            If Not PartDocs(0).InsertPages(n - 1, PartDocs(i), 0, ni, True) Then
                MsgBox "Cannot insert pages of" & vbLf & p & a(i), vbExclamation, "Canceled"
            End If

            n = n + ni

            PartDocs(i).Close
            Set PartDocs(i) = Nothing
        Else

            n = PartDocs(0).GetNumPages()
        End If
    Next
     
    If i > UBound(a) Then

        If Not PartDocs(0).Save(PDSaveFull, p & DestFile) Then
            MsgBox "Cannot save the resulting document" & vbLf & p & DestFile, vbExclamation, "Canceled"
        End If
    End If
     
exit_:
     

    If Err Then
        MsgBox Err.Description, vbCritical, "Error #" & Err.Number
    ElseIf i > UBound(a) Then
        MsgBox "The resulting file is created:" & vbLf & p & DestFile, vbInformation, "Done"
    End If
     

    If Not PartDocs(0) Is Nothing Then PartDocs(0).Close
    Set PartDocs(0) = Nothing
     
    
    AcroApp.Exit
    Set AcroApp = Nothing
     
    ForwardMerge3
     
End Sub

