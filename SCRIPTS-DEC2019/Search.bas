Attribute VB_Name = "Search"
Sub SelectBN()

    Dim selItem As Object
    Dim aMail As MailItem
    Dim aAttach As Attachment
 
    Dim a As Long
    Dim PDF, SubTxT, FN, EM, SN, BdNm, t As String
    Dim CbMrg, DSPLEML, SNDEML, SRCHPDF As Boolean
   
Call SaveAtt(strFolderpath, PDFCount, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15) '<- Functie om locatie van PDF bestanden te halen

Call AAPJEPUNTJE(BdNm, INTERN, SN, SubTxT, EM, t, WaardenEM, WaardenBdNm)                          '<- Functie om bedrijfsnaam uit emailadres te halen

If INTERN > 0 Then
CbEM = 1
Result = MsgBox(SN & vbLf & vbLf & "Retour sturen?", vbYesNo, "MAIL VAN COLLEGA")
If Result = 6 Then
Result = MsgBox(SN & vbLf & vbLf & "Vraag van collega?", vbYesNo, "MAIL VAN COLLEGA")
Else
GoTo doorgaan:
End If
If Result = 6 Then
Call Factuur_Retour(CbAB, PDF, FN, BdNm, CbMrg, CbEM, EM, ONDERWERP, AEAD, AEIO, AERC, AEEO, FMPD, WEAD, WEBT, BTNR, WEFD, WEFN, WEIB, WEKV, WEBD, OVRG, REDEN, DSPLEML, CbAFGH, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)
Else
CbEM = 2
'msgbox = ja op factuur terug
MsgBox "Factuur retour aan collega"
Call Factuur_Retour(CbAB, PDF, FN, BdNm, CbMrg, CbEM, EM, ONDERWERP, AEAD, AEIO, AERC, AEEO, FMPD, WEAD, WEBT, BTNR, WEFD, WEFN, WEIB, WEKV, WEBD, OVRG, REDEN, DSPLEML, CbAFGH, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)
End If
Exit Sub
End If

doorgaan:
           
           
           Bestandsnaam.FN = FN                     '<- FACTUURNUMMER | *Overbodig, is nog niet bekend*
           Bestandsnaam.SN = SN
            Bestandsnaam.t = t
           Bestandsnaam.EM = EM
         Bestandsnaam.BdNm = BdNm
       Bestandsnaam.SubTxT = SubTxT
  Bestandsnaam.WaardenBdNm = WaardenBdNm
    Bestandsnaam.WaardenEM = WaardenEM
     Bestandsnaam.PDFCount = PDFCount
       Bestandsnaam.INTERN = INTERN
Bestandsnaam.strFolderpath = strFolderpath
    Bestandsnaam.FNABCMPLT = FNABCMPLT

 Bestandsnaam.TxTBN1 = TxTBN1
 Bestandsnaam.TxTBN2 = TxTBN2
 Bestandsnaam.TxTBN3 = TxTBN3
 Bestandsnaam.TxTBN4 = TxTBN4
 Bestandsnaam.TxTBN5 = TxTBN5
 Bestandsnaam.TxTBN6 = TxTBN6
 Bestandsnaam.TxTBN7 = TxTBN7
 Bestandsnaam.TxTBN8 = TxTBN8
 Bestandsnaam.TxTBN9 = TxTBN9
Bestandsnaam.TxTBN10 = TxTBN10
Bestandsnaam.TxTBN11 = TxTBN11
Bestandsnaam.TxTBN12 = TxTBN12
Bestandsnaam.TxTBN13 = TxTBN13
Bestandsnaam.TxTBN14 = TxTBN14
Bestandsnaam.TxTBN15 = TxTBN15

 Bestandsnaam.ObBN1 = False
 Bestandsnaam.ObBN2 = False
 Bestandsnaam.ObBN3 = False
 Bestandsnaam.ObBN4 = False
 Bestandsnaam.ObBN5 = False
 Bestandsnaam.ObBN6 = False
 Bestandsnaam.ObBN7 = False
 Bestandsnaam.ObBN8 = False
 Bestandsnaam.ObBN9 = False
Bestandsnaam.ObBN10 = False
Bestandsnaam.ObBN11 = False
Bestandsnaam.ObBN12 = False
Bestandsnaam.ObBN13 = False
Bestandsnaam.ObBN14 = False
Bestandsnaam.ObBN15 = False

If PDFCount > 1 Then

'Bestandsnaam.DSPLEML = ""
'Bestandsnaam.CbMrg = ""

Bestandsnaam.MultiPage1.Value = 0

Bestandsnaam.Show
Else
CbMrg = False
PDF = TxTBN1
'PDF = PDFName(1)
Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)
End If

End Sub
Sub SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

    Dim searchString() As Variant
    Dim PDF_path As String
    Dim appObj As Object, AVDocObj As Object
      
    
    'SRCHPDF = True
        
   
    Dim objOL As Outlook.Application
       
    Dim x As Integer
    
    Dim searchStringI As String
    Dim searchStringII As String
    Dim searchStringIII As String
    Dim searchStringIV As String
    Dim searchStringV As String
    Dim searchStringVI As String
    Dim searchStringVII As String
    Dim searchStringVIII As String
    Dim searchStringIX As String
    Dim searchStringX As String
    Dim searchStringXI As String
    Dim searchStringXII As String
 
   searchStringI = "Gem"
   If Not PDF = "" Then
  searchStringII = Left(PDF, Len(PDF) - 4)
  End If
 searchStringIII = "Datu"
  searchStringIV = "Ink"
   searchStringV = "Rout"
  searchStringVI = "%"
 searchStringVII = "IBAN"
searchStringVIII = "B0"
  searchStringIX = "KvK"
   searchStringX = Left(BdNm, 3)
  searchStringXI = ""
 searchStringXII = ""

    Waarden.CbIO = True
  Waarden.CbAEAD = True
    Waarden.CbFN = True
    Waarden.CbFD = True
  Waarden.CbWEAD = True
  Waarden.CbWEBT = True
  Waarden.CbIBAN = True
  Waarden.CbBTNR = True
  Waarden.CbWEKV = True
   Waarden.CbPDF = True
   
   Waarden.CbARC = False
Waarden.CbANDERS = False
    Waarden.CbAB = False

    Waarden.CbHA = False
    Waarden.CbGB = False
   Waarden.CbINK = False
   Waarden.CbONb = False
  
      Waarden.RC = ""
    Waarden.OVRG = ""

  If PDFCount = 0 Then
      SRCHPDF = False
Waarden.CbPDF = False
Waarden.CbONb = True
 Waarden.OVRG = "geen pdf-factuur in bijlage mail"
           FN = "onbekend"
MsgBox "Geen PDF in mail"
  Else
     PDF_path = strFolderpath & PDF
     FN = Left(PDF, Len(PDF) - 4)        '<- FACTUURNUMMER | FACTUURNUMMER UIT PDFNAAM
   Waarden.FN = FN
  End If

If SRCHPDF = True Then

    If Dir(PDF_path) = "" Then
        MsgBox "File not found..."
        Exit Sub
    End If
   
    On Error Resume Next

    Set olapp = GetObject(, "Outlook.Application")
    Set appObj = CreateObject("AcroExch.App")
    
    If Err.Number <> 0 Then
        MsgBox "Error in creating the Adobe Application object..."
        Set appObj = Nothing
        Exit Sub
    End If
    
    Set AVDocObj = CreateObject("AcroExch.AVDoc")

    If Err.Number <> 0 Then
        MsgBox "Error in creating the AVDoc object..."
        Set AVDocObj = Nothing
          Set appObj = Nothing
        Exit Sub
    End If
    
    On Error GoTo 0
  
If AVDocObj.Open(PDF_path, "") = True Then
        
    AVDocObj.BringToFront

If AVDocObj.findtext(searchStringI, False, False, True) = False Then
        YesOrNoAnswerToMessageBox = MsgBox(searchStringI & " correct op factuur vermeld?", vbYesNo, "Tenaamstelling")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringI = False
                Else
                searchStringI = True
            End If
Else
        YesOrNoAnswerToMessageBox = MsgBox(searchStringI & " correct op factuur vermeld?", vbYesNo, "Tenaamstelling")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringI = False
                Else
                searchStringI = True
            End If
End If

    AVDocObj.BringToFront

If AVDocObj.findtext(searchStringII, False, False, True) = False Then
        
        FN = InputBox(searchStringII & " correct op factuur vermeld?", "Factuurnummer", FN)
            Select Case StrPtr(FN)
            Case 0
                searchStringII = False
            Case Else
                searchStringII = True
            End Select
Else
        FN = InputBox(searchStringII & " correct op factuur vermeld?", "Factuurnummer", FN)
            Select Case StrPtr(FN)
            Case 0
                searchStringII = False
            Case Else
                searchStringII = True
            End Select
End If

    AVDocObj.BringToFront

If AVDocObj.findtext(searchStringIII, False, False, True) = False Then
        YesOrNoAnswerToMessageBox = MsgBox(searchStringIII & " correct op factuur vermeld?", vbYesNo, "Factuurdatum")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringIII = False
                Else
                searchStringIII = True
            End If
Else
        YesOrNoAnswerToMessageBox = MsgBox(searchStringIII & " correct op factuur vermeld?", vbYesNo, "Factuurdatum")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringIII = False
                Else
                searchStringIII = True
            End If
End If

    AVDocObj.BringToFront

If AVDocObj.findtext(searchStringIV, False, False, True) = False Then
        YesOrNoAnswerToMessageBox = MsgBox(searchStringIV & " correct op factuur vermeld?", vbYesNo, "Inkoopordernummer")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringIV = False
                Else
                searchStringIV = True
            End If
Else
        YesOrNoAnswerToMessageBox = MsgBox(searchStringIV & " correct op factuur vermeld?", vbYesNo, "Inkoopordernummer")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringIV = False
                Else
                searchStringIV = True
            End If
End If

    AVDocObj.BringToFront

If AVDocObj.findtext(searchStringV, False, False, True) = False Then
        YesOrNoAnswerToMessageBox = MsgBox(searchStringV & " correct op factuur vermeld?", vbYesNo, "Routecode")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringV = False
                Else
                searchStringV = True
            End If
Else
        YesOrNoAnswerToMessageBox = MsgBox(searchStringV & " correct op factuur vermeld?", vbYesNo, "Routecode")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringV = False
                Else
                searchStringV = True
            End If
End If

    AVDocObj.BringToFront
            
If AVDocObj.findtext(searchStringX, False, False, True) = False Then
        YesOrNoAnswerToMessageBox = MsgBox(searchStringX & " correct op factuur vermeld?", vbYesNo, "NAW leverancier")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringX = False
                Else
                searchStringX = True
            End If
Else
        YesOrNoAnswerToMessageBox = MsgBox(searchStringX & " correct op factuur vermeld?", vbYesNo, "NAW leverancier")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringX = False
                Else
                searchStringX = True
            End If
End If

    AVDocObj.BringToFront

If AVDocObj.findtext(searchStringVI, False, False, True) = False Then
        YesOrNoAnswerToMessageBox = MsgBox(searchStringVI & " correct op factuur vermeld?", vbYesNo, "BTW-tarief")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringVI = False
                Else
                searchStringVI = True
            End If
Else
        YesOrNoAnswerToMessageBox = MsgBox(searchStringVI & " correct op factuur vermeld?", vbYesNo, "BTW-tarief")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringVI = False
                Else
                searchStringVI = True
            End If
End If

    AVDocObj.BringToFront

If AVDocObj.findtext(searchStringVII, False, False, True) = False Then
        YesOrNoAnswerToMessageBox = MsgBox(searchStringVII & " correct op factuur vermeld?", vbYesNo, "IBAN")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringVII = False
                Else
                searchStringVII = True
            End If
Else
        YesOrNoAnswerToMessageBox = MsgBox(searchStringVII & " correct op factuur vermeld?", vbYesNo, "IBAN")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringVII = False
                Else
                searchStringVII = True
            End If
End If

    AVDocObj.BringToFront

If AVDocObj.findtext(searchStringVIII, False, False, True) = False Then
        YesOrNoAnswerToMessageBox = MsgBox(searchStringVIII & " correct op factuur vermeld?", vbYesNo, "BTW-nummer")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringVIII = False
                Else
                searchStringVIII = True
            End If
Else
        YesOrNoAnswerToMessageBox = MsgBox(searchStringVIII & " correct op factuur vermeld?", vbYesNo, "BTW-nummer")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringVIII = False
                Else
                searchStringVIII = True
            End If
End If
        
    AVDocObj.BringToFront
            
If AVDocObj.findtext(searchStringIX, False, False, True) = False Then
        YesOrNoAnswerToMessageBox = MsgBox(searchStringIX & " correct op factuur vermeld?", vbYesNo, "KvK nummer")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringIX = False
                Else
                searchStringIX = True
            End If
Else
        YesOrNoAnswerToMessageBox = MsgBox(searchStringIX & " correct op factuur vermeld?", vbYesNo, "KvK nummer")
                If YesOrNoAnswerToMessageBox = vbNo Then
                searchStringIX = False
                Else
                searchStringIX = True
            End If
End If

End If
        AVDocObj.BringToFront

        SendKeys "^w"
        
        Set AVDocObj = Nothing
        Set appObj = Nothing

     Waarden.CbIO = True
If searchStringIV = False And searchStringV = False Then
     Waarden.CbIO = False
End If
    
      Waarden.CbAEAD = True
    If searchStringI = False Then
      Waarden.CbAEAD = False
        End If
        
            Waarden.CbFN = True
       If searchStringII = False Then
            Waarden.CbFN = False
            Waarden.CbONb = True
            FN = "onbekend"
            End If

                Waarden.CbFD = True
          If searchStringIII = False Then
                Waarden.CbFD = False
                End If

                    Waarden.CbWEBT = True
                 If searchStringVI = False Then
                    Waarden.CbWEBT = False
                    End If

                        Waarden.CbIBAN = True
                    If searchStringVII = False Then
                        Waarden.CbIBAN = False
                        End If

                            Waarden.CbBTNR = True
                       If searchStringVIII = False Then
                            Waarden.CbBTNR = False
                            End If

                                Waarden.CbWEKV = True
                             If searchStringIX = False Then
                                Waarden.CbWEKV = False
                                End If
                                
                                    Waarden.CbWEAD = True
                                 If searchStringX = False Then
                                    Waarden.CbWEAD = False
                                    End If
                                
                               
End If
 
  If Not DFA2 = "" Then
 Waarden.DFA2 = DFA2
  End If
 
  If Not DFA3 = "" Then
 Waarden.DFA3 = DFA3
  End If
 
  If Not DFA4 = "" Then
 Waarden.DFA4 = DFA4
  End If
 
      Waarden.PDF = PDF
   Waarden.SubTxT = SubTxT
       Waarden.EM = EM
     Waarden.CbEM = False
        Waarden.t = t
       Waarden.SN = SN
   Waarden.INTERN = INTERN
     Waarden.BdNm = BdNm
       Waarden.FN = FN
     
      Waarden.FN1 = FN
Waarden.FNABCMPLT = FNABCMPLT
 
  If Not CbMrg = "" Then
 Waarden.CbMrg = CbMrg
  End If
 
  If Not CbAFGH = "" Then
 Waarden.CbAFGH = CbAFGH
  End If
 
  If Not SRCHPDF = "" Then
 Waarden.SRCHPDF = SRCHPDF
  End If
 
  If Not DSPLEML = "" Then
 Waarden.DSPLEML = DSPLEML
  End If

Waarden.EM.BackColor = vbWhite
Waarden.BdNm.BackColor = vbWhite

If WaardenEM = True Then Waarden.EM.BackColor = vbRed

If WaardenBdNm = True Then Waarden.BdNm.BackColor = vbRed

        
Waarden.Show
           
End Sub
