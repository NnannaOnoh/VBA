Attribute VB_Name = "PDF"
Sub MergeToSend()
     
    Dim oExplorer As Outlook.Explorer
    Dim oMail As MailItem
    Set oExplorer = Application.ActiveExplorer
    Dim itm As Object
    
    SaveAtt
    
    SaveToMerge
End Sub
Sub ForwardMerge3()

    Dim oExplorer As Outlook.Explorer
    Dim oMail As MailItem
    Set oExplorer = Application.ActiveExplorer
    Dim itm As Object
    
    strUser = Left(Environ("USERNAME"), 3)

    Set oMail = oExplorer.Selection.Item(1).Forward
    
     Path = "H:\Mijn Documenten\merge\pdf\OLAttachments\"
     Ext = ".pdf"
    
    FN = InputBox("Factuurnummer")
    Select Case StrPtr(FN)
        Case 0
    Call KillAll
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
    On Error GoTo Release
     
    If oExplorer.Selection.Item(1).Class = olMail Then
        
        'oMail.HTMLBody = "Custom Text.<p> <img src=""custom image link""" _
        & " title=""D"" alt=""D"" name=""D"" border=""0"" id=""D""/>" _
        & vbCrLf & oMail.HTMLBody
        oMail.SentOnBehalfOfName = "facturen@amsterdam.nl"
        oMail.Recipients.Add "srvc47ACAM@amsterdam.nl"
        oMail.Recipients.Item(1).Resolve
        
        Do Until oMail.Attachments.Count = 0
            oMail.Attachments.Remove (1)
        Loop

        oMail.Attachments.Add (Path & FN & "(M)" & ".pdf")
        oMail.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
        oMail.Subject = oMail.Subject & " " & FN
        oMail.HTMLBody = oMail.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser
        'If omail.Recipients.item(1).Resolved Then
            'omail.Display
            'oMail.Save
            oMail.Display
        'Else
        '    MsgBox "Could not resolve " & omail.Recipients.item(1).Address
        'End If
    Else
        MsgBox ("Not a mail item")
    End If
Release:
    Set oMail = Nothing
    Set oExplorer = Nothing
    
'    Application.Wait Second(Now) + 15
    
    'Kill (Path & "*.jpg")
        
    Call KillAll
       
    Call Afgehandeld
    
End Sub
Sub KillAll()

Kill ("H:\Mijn Documenten\merge\pdf\OLAttachments\*.*")

End Sub
Sub SplitPerPDFCompleet()

    Dim oExplorer As Outlook.Explorer
    Dim oMail As MailItem
    Set oExplorer = Application.ActiveExplorer
    Dim itm As Object
    
    SaveAtt

    Set oMail = oExplorer.Selection.Item(1).Forward
    
     Path = "H:\Mijn Documenten\merge\pdf\OLAttachments\"
     Ext = ".pdf"
     
BN = InputBox("Bestandsnaam")
Select Case StrPtr(BN)
    Case 0
    Call KillAll
    MsgBox ("Geannuleerd")
        Exit Sub
    Case Else
End Select
    
    FN = InputBox("Factuurnummer")
    Select Case StrPtr(FN)
        Case 0
        Call KillAll
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
    On Error GoTo Release
     
    If oExplorer.Selection.Item(1).Class = olMail Then
        
        'oMail.HTMLBody = "Custom Text.<p> <img src=""custom image link""" _
        & " title=""D"" alt=""D"" name=""D"" border=""0"" id=""D""/>" _
        & vbCrLf & oMail.HTMLBody
        oMail.SentOnBehalfOfName = "facturen@amsterdam.nl"
        oMail.Recipients.Add "srvc47ACAM@amsterdam.nl"
        oMail.Recipients.Item(1).Resolve
        
        Do Until oMail.Attachments.Count = 0
            oMail.Attachments.Remove (1)
        Loop

        oMail.Attachments.Add (Path & BN & ".pdf")
        oMail.Subject = oMail.Subject & " " & FN
        'If omail.Recipients.item(1).Resolved Then
            'omail.Display
            'oMail.Save
            oMail.Send
        'Else
        '    MsgBox "Could not resolve " & omail.Recipients.item(1).Address
        'End If
    Else
        MsgBox ("Not a mail item")
    End If
Release:
    Set oMail = Nothing
    Set oExplorer = Nothing
    
'    Application.Wait Second(Now) + 15
    
    'Kill (Path & "*.jpg")
    Call KillAll
       
    'Call Afgehandeld
    
End Sub
Sub SaveToMerge()
     
    Dim DestFile As String
    
    FN = InputBox("Factuurnummer")
    Select Case StrPtr(FN)
        Case 0
    Call KillAll
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
            Call KillAll
            MsgBox "File not found" & vbLf & p & a(i), vbExclamation, "Canceled"
            Exit For
        End If

        Set PartDocs(i) = CreateObject("AcroExch.PDDoc")
        PartDocs(i).Open p & Trim(a(i))
        If i Then

            ni = PartDocs(i).GetNumPages()
            If Not PartDocs(0).InsertPages(n - 1, PartDocs(i), 0, ni, True) Then
            Call KillAll
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
        Call KillAll
            MsgBox "Cannot save the resulting document" & vbLf & p & DestFile, vbExclamation, "Canceled"
        End If
    End If
     
exit_:
     

    If Err Then
        MsgBox Err.Description, vbCritical, "Error #" & Err.Number & " Comma in bestandsnaam!"
        MsgBox "mail niet verzonden, annuleer opdracht!"
    ElseIf i > UBound(a) Then
           MsgBox "The resulting file is created:" & vbLf & p & DestFile, vbInformation, "Done"
    End If
     

    If Not PartDocs(0) Is Nothing Then PartDocs(0).Close
    Set PartDocs(0) = Nothing
     
    
    AcroApp.Exit
    Set AcroApp = Nothing
     
    ForwardMerge3
     
End Sub
Sub SplitPerPDF_Compleet()

    Dim oExplorer As Outlook.Explorer
    Dim oMail As MailItem
    Set oExplorer = Application.ActiveExplorer
    Dim itm As Object
    
    SaveAtt

    Set oMail = oExplorer.Selection.Item(1).Forward
    
     Path = "H:\Mijn Documenten\merge\pdf\OLAttachments\"
     Ext = ".pdf"
     
BN = InputBox("Bestandsnaam")
Select Case StrPtr(BN)
    Case 0
    Call KillAll
    MsgBox ("Geannuleerd")
        Exit Sub
    Case Else
End Select
    
    FN = InputBox("Factuurnummer")
    Select Case StrPtr(FN)
        Case 0
        Call KillAll
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
    On Error GoTo Release
     
    If oExplorer.Selection.Item(1).Class = olMail Then
        
        'oMail.HTMLBody = "Custom Text.<p> <img src=""custom image link""" _
        & " title=""D"" alt=""D"" name=""D"" border=""0"" id=""D""/>" _
        & vbCrLf & oMail.HTMLBody
        oMail.SentOnBehalfOfName = "facturen@amsterdam.nl"
        oMail.Recipients.Add "srvc47ACAM@amsterdam.nl"
        oMail.Recipients.Item(1).Resolve
        
        Do Until oMail.Attachments.Count = 0
            oMail.Attachments.Remove (1)
        Loop

        oMail.Attachments.Add (Path & BN & ".pdf")
        oMail.Subject = oMail.Subject & " " & FN
        'If omail.Recipients.item(1).Resolved Then
            oMail.Display
            'oMail.Save
            'oMail.Send
        'Else
        '    MsgBox "Could not resolve " & omail.Recipients.item(1).Address
        'End If
    Else
        MsgBox ("Not a mail item")
    End If
Release:
    Set oMail = Nothing
    Set oExplorer = Nothing
    
'    Application.Wait Second(Now) + 15
    
    'Kill (Path & "*.jpg")
    Call KillAll
       
    'Call Afgehandeld
    
End Sub
Sub SplitPerPDF_Routecode()

    Dim oExplorer As Outlook.Explorer
    Dim oMail As MailItem
    Set oExplorer = Application.ActiveExplorer
    Dim itm As Object
    
    SaveAtt

    Set oMail = oExplorer.Selection.Item(1).Reply
    
     Path = "H:\Mijn Documenten\merge\pdf\OLAttachments\"
     Ext = ".pdf"
     
BN = InputBox("Bestandsnaam")
Select Case StrPtr(BN)
    Case 0
    Call KillAll
    MsgBox ("Geannuleerd")
        Exit Sub
    Case Else
End Select
    
FN = InputBox("Factuurnummer")
Select Case StrPtr(FN)
    Case 0
    Call KillAll
    MsgBox ("Geannuleerd")
        Exit Sub
    Case Else
End Select
    
 b = InputBox("Bedrijfsnaam")
Select Case StrPtr(FN)
        Case 0
        Call KillAll
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
End Select
    
    'On Error GoTo Release
     
    If oExplorer.Selection.Item(1).Class = olMail Then
        
        'oMail.HTMLBody = "Custom Text.<p> <img src=""custom image link""" _
        & " title=""D"" alt=""D"" name=""D"" border=""0"" id=""D""/>" _
        & vbCrLf & oMail.HTMLBody
        oMail.SentOnBehalfOfName = "facturen@amsterdam.nl"
        'oMail.Recipients.Add ""
        oMail.Recipients.Item(1).Resolve
        
        Do Until oMail.Attachments.Count = 0
            oMail.Attachments.Remove (1)
        Loop

        oMail.Attachments.Add (Path & BN & ".pdf")
        oMail.Subject = "Teruggestuurd/Factuur " & FN & "/" & b & "/AE"
        oMail.HTMLBody = "<font-size=""10.5"" face=""Corbel"" color=""black"">" _
        & "Geachte relatie, <br><br>" _
        & "Onlangs heeft u ons bijgevoegde factuur gestuurd. De factuur bevat helaas niet alle gegevens die wij nodig hebben om te kunnen betalen. Wij nemen deze factuur daarom nu niet in behandeling." _
        & "<br><br>Vanwege de vernieuwing en verbetering van onze financiële administratie vragen we onze leveranciers sinds enige tijd om aanpassing van de facturen. Mogelijk heeft ons verzoek u niet bereikt." _
        & "<br><br>Wij vragen u de factuur te controleren op de aangegeven onderdelen:<br><br>" _
        & "<li><b>Routecode/kostenplaats; </b><u><i>geen of geen geldige routecode/inkoopordernummer op factuur vermeld</u></i></li>" _
        & "<br><br>Wilt u uw factuur waar nodig aanpassen?<b><u> Als u geen routecode of inkoopordernummer heeft, kunt u contact opnemen met uw opdrachtgever of contactpersoon binnen de gemeente Amsterdam. </b></u>" _
        & "<br>" _
        & "<br>U kunt de aangepaste factuur (inclusief eventuele bijlagen in hetzelfde PDF-bestand) sturen naar <a href=mailto:facturen@amsterdam.nl>facturen@amsterdam.nl</a>." _
        & "<br><br>" _
        & "Wij sturen een bericht naar aanleiding van elke onvolledige factuur die u ons stuurt. Het kan dus zijn dat u meerdere berichten van ons ontvangt.  We kunnen ons voorstellen dat het voor u vervelend is om deze aanpassingen te doen. Het helpt ons om u sneller te kunnen betalen. We bedanken u daarom voor uw hulp." _
        & "<br>" _
        & "<br>" _
        & "Als u vragen heeft  kunt u contact met ons opnemen via <a href=mailto:crediteurenadministratie@amsterdam.nl>crediteurenadministratie@amsterdam.nl</a>.<br>" _
        & "<br>" _
        & "<br>" _
        & "Met vriendelijke groet, " _
        & "<br>" _
        & "<br>" _
        & "<br>" _
        & "Crediteurenadministratie<br>" _
        & "<font size=""3"" face=""Corbel"" color=""red""><b>Gemeente Amsterdam</b><br>" _
        & "<br></font>" & oMail.HTMLBody _
        
        'If omail.Recipients.item(1).Resolved Then
            'omail.Display
            'oMail.Save
            oMail.Display
        'Else
        '    MsgBox "Could not resolve " & omail.Recipients.item(1).Address
        'End If
    Else
        MsgBox ("Not a mail item")
    End If
Release:
    Set oMail = Nothing
    Set oExplorer = Nothing
    
'    Application.Wait Second(Now) + 15
    
    'Kill (Path & "*.jpg")
    Call KillAll
       
    'Call Retour
    
End Sub
Sub watermerkkkkk()

Dim Path As String
Dim BN As String


    Dim oExplorer As Outlook.Explorer
    Dim oMail As MailItem
    Set oExplorer = Application.ActiveExplorer
    Dim itm As Object
    
    SaveAtt

BN = InputBox("Bestandsnaam")
Select Case StrPtr(BN)
    Case 0
    Call KillAll
    MsgBox ("Geannuleerd")
        Exit Sub
    Case Else
End Select

'Application.DisplayAlerts = False


File = "H:\Mijn Documenten\merge\pdf\OLAttachments\File.PDF"

Set app = CreateObject("Acroexch.app")
app.Hide
Set avDoc = CreateObject("AcroExch.avDoc")
Set AForm = CreateObject("AFormAut.App")

If avDoc.Open(File, "") Then
Set PDDoc = avDoc.GetPDDoc()
    Set jso = PDDoc.GetJSObject
       
    Ex = "  //  set Date, filename and PageNo as footer " & vbLf _
      & "  var Box2Width = 100  " & vbLf _
      & "  for (var p = 0; p < this.numPages; p++)   " & vbLf _
      & "   {   " & vbLf _
      & "    var aRect = this.getPageBox(""Crop"",p);  " & vbLf _
      & "    var TotWidth = aRect[2] - aRect[0]  " & vbLf _
      & "     {  var bStart=(TotWidth/1)-(Box2Width/4)  " & vbLf _
      & "         var bEnd=((TotWidth/2)+(Box2Width/2))  " & vbLf _
      & "         var fp = this.addField(String(""xftPage""+p+1), ""text"", p, [bStart,25,bEnd,60]);   " & vbLf _
      & "         fp.value = ""GECONTROLEERD DOOR INPUT"" + util.printd("" dd/mm/yyyy"", new Date());  " & vbLf _
      & "         fp.textSize=8;  fp.color=BLUE; fp.readonly = true;  " & vbLf _
      & "         fp.alignment=""center"";  " & vbLf _
      & "     }  " & vbLf _
      & "   }  "
      



       AForm.Fields.ExecuteThisJavaScript Ex
       
PDDoc.Save PDSaveIncremental, FileNm

PDDoc.Close

app.Show

    MsgBox ("Done")
    

End If

Set avDoc = Nothing
Set app = Nothing
Set PDDoc = Nothing

End Sub




