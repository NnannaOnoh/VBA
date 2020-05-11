Attribute VB_Name = "Routecode"

Sub GeenRoutecode()
    Dim rpl As Outlook.MailItem
    Dim itm As Object
    Dim verzonden As Integer
         
    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set rpl = itm.Reply
        CopyAttachments itm, rpl
        
        i = InputBox("Factuurnummer")
    Select Case StrPtr(i)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
        b = InputBox("Bedrijfsnaam")
    Select Case StrPtr(b)
        Case 0
        MsgBox ("Geannuleerd")
            Exit Sub
        Case Else
    End Select
    
        rpl.SentOnBehalfOfName = "facturen@amsterdam.nl"
        rpl.Recipients.Item(1).Resolve
        'rpl.Recipients.Add ""
        
        rpl.Subject = "Teruggestuurd/Factuur " & i & "/" & b & "/AE"
        
        rpl.HTMLBody = "<font size=""3"" face=""Corbel"" color=""black"">" _
        & "Geachte relatie, <br><br>" _
        & "Onlangs heeft u ons bijgevoegde factuur gestuurd. De factuur bevat helaas niet alle gegevens die wij nodig hebben om te kunnen betalen. Wij nemen deze factuur daarom nu niet in behandeling." _
        & "<br><br>Vanwege de vernieuwing en verbetering van onze financiële administratie vragen we onze leveranciers sinds enige tijd om aanpassing van de facturen. Mogelijk heeft ons verzoek u niet bereikt." _
        & "<br><br>Wij vragen u de factuur te controleren op de aangegeven onderdelen:" _
        & "<br><br>" _
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
        & "<br></font>" & rpl.HTMLBody _
        
        rpl.Attachments.Add ("G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\archief\flyer_1_administratie_a5_digitaal.pdf")

        rpl.Display
        
    End If
 
    Set rpl = Nothing
    Set itm = Nothing
    
        Call Retour

End Sub


