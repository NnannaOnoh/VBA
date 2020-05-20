Attribute VB_Name = "MERGEN"
Sub MERGE(DFA, PDF, FN, CbMrg, CbAFGH, DSPLEML, SNDEML, CbARC, CbINK, EM, t, SubTxT, SN, REC, RC)
     
    Dim DestFile As String
    
    DestFile = PDF
    
    Dim MyPath As String, MyFiles As String, ToPath As String
    
    Dim a() As String, i As Long, f As String
     

    MyPath = "H:\Mijn Documenten\merge\pdf\OLAttachments\"


    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    ReDim a(1 To 2 ^ 14)
    
            f = Dir(MyPath & PDF)
            i = i + 1
            a(i) = f
            f = Dir()
            
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

        Call MergePDFs(MyPath, MyFiles, DestFile, DFA, PDF, FN, CbAFGH, DSPLEML, SNDEML, CbARC, CbINK, CbMrg, EM, t, SubTxT, SN, REC, RC)
    Else
        MsgBox "No PDF files found in" & vbLf & MyPath, vbExclamation, "Canceled"
    End If

End Sub
Sub MergePDFs(MyPath, MyFiles, DestFile, DFA, PDF, FN, CbAFGH, DSPLEML, SNDEML, CbARC, CbINK, CbMrg, EM, t, SubTxT, SN, REC, RC)
     
    Dim a As Variant, i As Long, N As Long, ni As Long, p As String
    Dim AcroApp As New Acrobat.AcroApp, PartDocs() As Acrobat.CAcroPDDoc
    
    DestFile = Left(DestFile, Len(DestFile) - 4) & "(M)" & ".pdf"
     
    If Right(MyPath, 1) = "\" Then p = MyPath Else p = MyPath & "\"
    a = Split(MyFiles, ",")
    ReDim PartDocs(0 To UBound(a))
     
    On Error GoTo exit_
    If Len(Dir(p & DestFile)) Then Kill p & DestFile
    For i = 0 To UBound(a)

        If Dir(p & Trim(a(i))) = "" Then
            Call KillAll
            MsgBox "File not found" & vbLf & p & a(i), vbExclamation, "Canceled"
            Exit For
        End If

        Set PartDocs(i) = CreateObject("AcroExch.PDDoc")
        PartDocs(i).Open p & Trim(a(i))
        If i Then

            ni = PartDocs(i).GetNumPages()
            If Not PartDocs(0).InsertPages(N - 1, PartDocs(i), 0, ni, True) Then
            Call KillAll
                MsgBox vbLf & p & a(i) & " is mogelijk beveiligd.", vbExclamation, "Canceled"
            End If

            N = N + ni

            PartDocs(i).Close
            Set PartDocs(i) = Nothing
        Else

            N = PartDocs(0).GetNumPages()
        End If
    Next
     
    If i > UBound(a) Then

        If Not PartDocs(0).Save(PDSaveFull, p & DestFile) Then
        Call KillAll
            MsgBox "Cannot save the resulting document" & vbLf & p & DestFile, vbExclamation, "Canceled"
        End If
    End If
     
exit_:

    If Err Then
    
        MsgBox Err.Description, vbCritical, "Error #" & Err.Number & " Comma in bestandsnaam!"
        MsgBox "mail niet verzonden, annuleer opdracht!"
        
        Call KillAll
    Exit Sub
        
    ElseIf i > UBound(a) Then
           MsgBox "The resulting file is created:" & vbLf & p & DestFile, vbInformation, "Done"
    End If
     

    If Not PartDocs(0) Is Nothing Then PartDocs(0).Close
    Set PartDocs(0) = Nothing
   
    AcroApp.Exit
    Set AcroApp = Nothing

     
If CbARC = True Then
Call RCALTERNATIEF(DFA, PDF, DestFile, FN, CbAFGH, DSPLEML, SNDEML, CbARC, CbMrg, EM, t, SubTxT, SN, REC, RC)
Exit Sub
End If

If CbINK = True Then
Call IOINACTIEF(DFA, PDF, DestFile, FN, CbAFGH, DSPLEML, SNDEML, CbARC, CbINK, CbMrg, EM, t, SubTxT, SN, REC, RC)
Exit Sub
End If

PDF = DestFile

Call Factuur_Compleet(PDF, DFA, FN, DSPLEML, CbARC, CbINK, CbAFGH, CbSendAll, CbAB, EM, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)
  
End Sub
Sub KillAll()

Dim WMI As Object, Process As Object, ProcessToKill As String, strPath As String

  strPath = "H:\Mijn Documenten\merge\pdf\OLAttachments"
  If Right(strPath, 1) <> "\" Then strPath = strPath & "\"

If UCase(Dir(strPath & "*.PDF")) = "" Then
Exit Sub
Else
On Error Resume Next
Kill strPath & "*.*"
End If

If UCase(Dir(strPath & "*.PDF")) = "" Then
Exit Sub
Else

ProcessToKill = "acrotray.exe"
Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_Process")
For Each Process In WMI
    If Process.Name = ProcessToKill Then
        Process.Terminate
    End If

Next

Kill strPath & "*.*"

If UCase(Dir(strPath & "*.PDF")) = "" Then Exit Sub

On Error Resume Next

ProcessToKill = "AcroRd32.exe"
Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_Process")
For Each Process In WMI
    If Process.Name = ProcessToKill Then
        Process.Terminate
    End If

Next

Kill strPath & "*.*"

If UCase(Dir(strPath & "*.PDF")) = "" Then Exit Sub

On Error Resume Next

ProcessToKill = "Acrobat.exe"
Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_Process")
For Each Process In WMI
    If Process.Name = ProcessToKill Then
        Process.Terminate
    End If

Next

Kill strPath & "*.*"

End If

If UCase(Dir(strPath & "*.PDF")) = "" Then
Exit Sub
Else
MsgBox "PDF bestanden in " & strPath & " Verwijder bestanden."
Shell "C:\WINDOWS\explorer.exe " & strPath & ", vbMinimizeFocus"
End If

End Sub
