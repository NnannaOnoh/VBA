Attribute VB_Name = "WATERMERK"
Sub watermeeeeeerr()

Dim Path As String, BN As String, File As String, FileExt As String

SaveAtt2

Set app = CreateObject("Acroexch.app")
app.Show
Set avDoc = CreateObject("AcroExch.avDoc")
Set AForm = CreateObject("AFormAut.App")

'Application.DisplayAlerts = False

BN = InputBox("Bestandsnaam")


FileExt = "*.pdf*"

If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
 
File = "H:\Mijn Documenten\merge\pdf\OLAttachments\watermerk\" & BN & ".pdf"
   
'Do While Len(File) > 0
   
'File = "H:\FOTRON FACTUREN_Deel28.pdf"

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
      & "         fp.textSize=8;  fp.color=RED; fp.readonly = true;  " & vbLf _
      & "         fp.alignment=""center"";  " & vbLf _
      & "     }  " & vbLf _
      & "   }  "
      



       AForm.Fields.ExecuteThisJavaScript Ex
       
PDDoc.Save PDSaveIncremental, FileNm
    
PDDoc.Close

        MsgBox ("Done")

End If

'Loop

'AcroApp.Exit

Set app = Nothing
Set avDoc = Nothing
Set AForm = Nothing
   
    MsgBox ("Done")
  
    
    
    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As attachment
    Dim Report As String
     
    For Each selItem In Application.ActiveExplorer.Selection
        If selItem.Class = olMail Then
            Set aMail = selItem
            For Each aAttach In aMail.Attachments
                Report = Report & GetAttachmentInfo(aAttach)
                Report = Report & "; "


        
    Next
  
    Dim fwd As Outlook.MailItem
    Dim itm As Object
    Dim strUser As String
 
    strUser = Left(Environ("USERNAME"), 3)
   
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set fwd = itm.Forward
        
        Do Until fwd.Attachments.Count = 0
            fwd.Attachments.Remove (1)
        Loop
        
        fwd.SentOnBehalfOfName = "Facturen@amsterdam.nl"
        fwd.Recipients.Add "srvc47ACAM@amsterdam.nl"
        
        'CopyAttachments itm, fwd
        fwd.Subject = fwd.Subject
        fwd.Attachments.Add "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\VHB.png", olByValue, 0
        fwd.Attachments.Add "H:\Mijn Documenten\merge\pdf\OLAttachments\watermerk\" & BN & ".pdf"
        fwd.HTMLBody = fwd.HTMLBody & "<br><br>_____________________________________________________&nbsp;" & "<br><img src='cid:VHB.png'" & "width='27' height='17'>" & Report _
                      & "<img src='cid:VHB.png'" & "width='27' height='17'>" & "<br><font size=""3"" face=""Corbel"" color=""white"">" _
                      & (Format(Now, "yyyy-mm-dd hh:mm:ss")) & " <b>Gemeente Amsterdam </b> " & strUser
        
        fwd.Display
        fwd.DeferredDeliveryTime = DateAdd("s", 25, Now)
        
    End If
     
    Set fwd = Nothing
    Set itm = Nothing
    
    'Kill "H:\Mijn Documenten\merge\pdf\OLAttachments\watermerk\" & BN & ".pdf"
    
    'Call Afgehandeld
    
    End If
    
    Next
    
End Sub
Sub HANDMATIG(Report, BN)

Dim Path As String, File As String, FileExt As String

SaveAtt2

Set app = CreateObject("Acroexch.app")
app.Show
Set avDoc = CreateObject("AcroExch.avDoc")
Set AForm = CreateObject("AFormAut.App")

'Application.DisplayAlerts = False


FileExt = "*.pdf*"

If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
 
File = "H:\Mijn Documenten\merge\pdf\OLAttachments\watermerk\" & BN & ".pdf"
   
'Do While Len(File) > 0
   
'File = "H:\FOTRON FACTUREN_Deel28.pdf"

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
      & "         fp.value = ""BETAALSTATUS: H"" + ""   "" + ""TEAM 1"" + ""   "" + util.printd("" dd/mm/yy"", new Date());  " & vbLf _
      & "         fp.textSize=8;  fp.color=RED; fp.readonly = true;  " & vbLf _
      & "         fp.alignment=""center"";  " & vbLf _
      & "     }  " & vbLf _
      & "   }  "
      



       AForm.Fields.ExecuteThisJavaScript Ex
       
PDDoc.Save PDSaveIncremental, FileNm
    
PDDoc.Close

End If

'Loop

'AcroApp.Exit

Set app = Nothing
Set avDoc = Nothing
Set AForm = Nothing
   
End Sub




