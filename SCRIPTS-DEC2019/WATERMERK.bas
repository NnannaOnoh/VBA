Attribute VB_Name = "WATERMERK"
Sub HANDMATIG()

Dim File, InitFolder As String
Set xlApp = Excel.Application

Set app = CreateObject("Acroexch.app")
Set AvDoc = CreateObject("AcroExch.avDoc")
Set AForm = CreateObject("AFormAut.App")

InitFolder = "H:\Mijn Documenten\merge\pdf\OLAttachments\"
   
    On Error Resume Next
    
    With xlApp.FileDialog(msoFileDialogFilePicker)
        .Show
        InitFolder = .InitialFolderName
        File = .SelectedItems(1)
    End With

If AvDoc.Open(File, "") Then
Set PDDoc = AvDoc.GetPDDoc()
    Set jso = PDDoc.GetJSObject
       
    Ex = "  //  set Date, filename and PageNo as header " & vbLf _
      & "  var Box2Width = 100  " & vbLf _
      & "  for (var p = 0; p < this.numPages; p++)   " & vbLf _
      & "   {   " & vbLf _
      & "    var aRect = this.getPageBox(""Crop"",p);  " & vbLf _
      & "    var TotWidth = aRect[2] - aRect[0]  " & vbLf _
      & "     {  var bStart=(TotWidth/3)-(Box2Width/1)  " & vbLf _
      & "         var bEnd=((TotWidth/1.30)+(Box2Width/1))  " & vbLf _
      & "         var fp = this.addField(String(""xftPage""+p+1), ""text"", p, [bStart,810,bEnd,830]);   " & vbLf _
      & "         fp.value = ""              |                                                                                         | "";  " & vbLf _
      & "         fp.borderStyle = border.s; fp.strokeColor = color.blue; fp.lineWidth = 2;  " & vbLf _
      & "         fp.textSize=12;  fp.textColor=color.red ; fp.readonly = FALSE;  " & vbLf _
      & "         fp.alignment=""center"";  " & vbLf _
      & "     }  " & vbLf _
      & "   }  "
      
       AForm.Fields.ExecuteThisJavaScript Ex
       
PDDoc.Save PDSaveIncremental, FileNM

app.Show

PDDoc.Close

End If

Set app = Nothing
Set AvDoc = Nothing
Set AForm = Nothing
Set xlApp = Nothing

End Sub
Sub HANDMATIG2(Report, BN)

Dim Path As String, File As String, FileExt As String

Set app = CreateObject("Acroexch.app")

Set AvDoc = CreateObject("AcroExch.avDoc")
Set AForm = CreateObject("AFormAut.App")

FileExt = "*.pdf*"

If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
 

File = "H:\Mijn Documenten\merge\pdf\OLAttachments\" & BN

FileNM = File

If AvDoc.Open(File, "") Then
Set PDDoc = AvDoc.GetPDDoc()
    Set jso = PDDoc.GetJSObject
       
    Ex = "  //  set Date, filename and PageNo as header " & vbLf _
      & "  var Box2Width = 100  " & vbLf _
      & "  for (var p = 0; p < this.numPages; p++)   " & vbLf _
      & "   {   " & vbLf _
      & "    var aRect = this.getPageBox(""Crop"",p);  " & vbLf _
      & "    var TotWidth = aRect[2] - aRect[0]  " & vbLf _
      & "     {  var bStart=(TotWidth/1)-(Box2Width/4)  " & vbLf _
      & "         var bEnd=((TotWidth/2.05)+(Box2Width/2))  " & vbLf _
      & "         var fp = this.addField(String(""xftPage""+p+1), ""text"", p, [bStart,810,bEnd,830]);   " & vbLf _
      & "         fp.value = ""CREDIT | TEAM 1 | "" + util.printd("" dd/mm/yy"", new Date());  " & vbLf _
      & "         fp.borderStyle = border.s; fp.strokeColor = color.green; fp.lineWidth = 2;  " & vbLf _
      & "         fp.textSize=12;  fp.textColor=color.green ; fp.readonly = true;  " & vbLf _
      & "         fp.alignment=""center"";  " & vbLf _
      & "     }  " & vbLf _
      & "   }  "
      
       AForm.Fields.ExecuteThisJavaScript Ex
       
app.Show

PDDoc.Save PDSaveIncremental, FileNM

End If

Set app = Nothing
Set AvDoc = Nothing
Set AForm = Nothing
   
End Sub
Sub CREDCHCK(Report, BN)

Dim Path As String, File As String, FileExt As String

Set app = CreateObject("Acroexch.app")

Set AvDoc = CreateObject("AcroExch.avDoc")
Set AForm = CreateObject("AFormAut.App")

FileExt = "*.pdf*"

If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
 

File = "H:\Mijn Documenten\merge\pdf\OLAttachments\" & BN

FileNM = File


If AvDoc.Open(File, "") Then
Set PDDoc = AvDoc.GetPDDoc()
    Set jso = PDDoc.GetJSObject
       
    Ex = "  //  set Date, filename and PageNo as header " & vbLf _
      & "  var Box2Width = 100  " & vbLf _
      & "  for (var p = 0; p < this.numPages; p++)   " & vbLf _
      & "   {   " & vbLf _
      & "    var aRect = this.getPageBox(""Crop"",p);  " & vbLf _
      & "    var TotWidth = aRect[2] - aRect[0]  " & vbLf _
      & "     {  var bStart=(TotWidth/3)-(Box2Width/1)  " & vbLf _
      & "         var bEnd=((TotWidth/1.30)+(Box2Width/1))  " & vbLf _
      & "         var fp = this.addField(String(""xftPage""+p+1), ""text"", p, [bStart,810,bEnd,830]);   " & vbLf _
      & "         fp.value = ""CHECK CR T1 | DEB.DOC:    |RC:    |GB:    | "";  " & vbLf _
      & "         fp.borderStyle = border.s; fp.strokeColor = color.blue; fp.lineWidth = 2;  " & vbLf _
      & "         fp.textSize=12;  fp.textColor=color.red ; fp.readonly = FALSE;  " & vbLf _
      & "         fp.alignment=""center"";  " & vbLf _
      & "     }  " & vbLf _
      & "   }  "
      
       AForm.Fields.ExecuteThisJavaScript Ex
       
PDDoc.Save PDSaveIncremental, FileNM

app.Show

End If

Set app = Nothing
Set AvDoc = Nothing
Set AForm = Nothing
   
End Sub
Sub RCALTERNATIEF(DFA, PDF, DestFile, FN, CbAFGH, DSPLEML, SNDEML, CbARC, CbMrg, EM, t, SubTxT, SN, REC, RC)

Dim Path As String, File As String, FileExt As String

Set app = CreateObject("Acroexch.app")
Set AvDoc = CreateObject("AcroExch.avDoc")
Set AForm = CreateObject("AFormAut.App")

If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If

If CbMrg = True Then PDF = DestFile

File = "H:\Mijn Documenten\merge\pdf\OLAttachments\" & PDF

FileNM = File

If AvDoc.Open(File, "") Then
AvDoc.BringToFront
Set PDDoc = AvDoc.GetPDDoc()
    Set jso = PDDoc.GetJSObject
       
    Ex = "  //  set Date, filename and PageNo as header " & vbLf _
      & "  var Box2Width = 100  " & vbLf _
      & "  for (var p = 0; p < this.numPages; p++)   " & vbLf _
      & "   {   " & vbLf _
      & "    var aRect = this.getPageBox(""Crop"",p);  " & vbLf _
      & "    var TotWidth = aRect[2] - aRect[0]  " & vbLf _
      & "     {  var bStart=(TotWidth/3)-(Box2Width/1)  " & vbLf _
      & "         var bEnd=((TotWidth/1.30)+(Box2Width/1))  " & vbLf _
      & "         var fp = this.addField(String(""xftPage""+p+1), ""text"", p, [bStart,810,bEnd,830]);   " & vbLf _
      & "         fp.value = ""RC INACTIEF | ALTERNATIEVE RC:   | "";  " & vbLf _
      & "         fp.borderStyle = border.s; fp.strokeColor = color.blue; fp.lineWidth = 2;  " & vbLf _
      & "         fp.textSize=12;  fp.textColor=color.red ; fp.readonly = FALSE;  " & vbLf _
      & "         fp.alignment=""center"";  " & vbLf _
      & "     }  " & vbLf _
      & "   }  "
      
       AForm.Fields.ExecuteThisJavaScript Ex
       
PDDoc.Save PDSaveIncremental, FileNM

app.Show

End If

Set app = Nothing
Set AvDoc = Nothing
Set AForm = Nothing

Call Factuur_Compleet(PDF, DFA, FN, DSPLEML, CbARC, CbINK, CbAFGH, CbSendAll, CbAB, EM, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)
   
End Sub
Sub IOINACTIEF(DFA, PDF, DestFile, FN, CbAFGH, DSPLEML, SNDEML, CbARC, CbINK, CbMrg, EM, t, SubTxT, SN, REC, RC)

Dim Path As String, File As String, FileExt As String

Set app = CreateObject("Acroexch.app")
Set AvDoc = CreateObject("AcroExch.avDoc")
Set AForm = CreateObject("AFormAut.App")

If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If

If CbMrg = True Then PDF = DestFile

File = "H:\Mijn Documenten\merge\pdf\OLAttachments\" & PDF

FileNM = File

If AvDoc.Open(File, "") Then
Set PDDoc = AvDoc.GetPDDoc()
    Set jso = PDDoc.GetJSObject
       
    Ex = "  //  set Date, filename and PageNo as header " & vbLf _
      & "  var Box2Width = 100  " & vbLf _
      & "  for (var p = 0; p < this.numPages; p++)   " & vbLf _
      & "   {   " & vbLf _
      & "    var aRect = this.getPageBox(""Crop"",p);  " & vbLf _
      & "    var TotWidth = aRect[2] - aRect[0]  " & vbLf _
      & "     {  var bStart=(TotWidth/3)-(Box2Width/1)  " & vbLf _
      & "         var bEnd=((TotWidth/1.30)+(Box2Width/1))  " & vbLf _
      & "         var fp = this.addField(String(""xftPage""+p+1), ""text"", p, [bStart,810,bEnd,830]);   " & vbLf _
      & "         fp.value = ""IO INACTIEF:    | RC:    | "";  " & vbLf _
      & "         fp.borderStyle = border.s; fp.strokeColor = color.blue; fp.lineWidth = 2;  " & vbLf _
      & "         fp.textSize=12;  fp.textColor=color.red ; fp.readonly = FALSE;  " & vbLf _
      & "         fp.alignment=""center"";  " & vbLf _
      & "     }  " & vbLf _
      & "   }  "
      
       AForm.Fields.ExecuteThisJavaScript Ex
       
PDDoc.Save PDSaveIncremental, FileNM

app.Show

End If

Set app = Nothing
Set AvDoc = Nothing
Set AForm = Nothing

Call Factuur_Compleet(PDF, DFA, FN, DSPLEML, CbARC, CbINK, CbAFGH, CbSendAll, CbAB, EM, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)
  
End Sub
