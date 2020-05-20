VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Bestandsnaam 
   Caption         =   "BESTAND"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   OleObjectBlob   =   "Bestandsnaam.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Bestandsnaam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SENDALL_Click()

If DFA = "" Then DFA = "srvc18VR@amsterdam.nl"

FN = Left(TxTBN1, Len(TxTBN1) - 4)
If Not TxTBN2 = "" Then FN = FN & " | " & Left(TxTBN2, Len(TxTBN2) - 4)
If Not TxTBN3 = "" Then FN = FN & " | " & Left(TxTBN3, Len(TxTBN3) - 4)
If Not TxTBN4 = "" Then FN = FN & " | " & Left(TxTBN4, Len(TxTBN4) - 4)
If Not TxTBN5 = "" Then FN = FN & " | " & Left(TxTBN5, Len(TxTBN5) - 4)
If Not TxTBN6 = "" Then FN = FN & " | " & Left(TxTBN6, Len(TxTBN6) - 4)
If Not TxTBN7 = "" Then FN = FN & " | " & Left(TxTBN7, Len(TxTBN7) - 4)
If Not TxTBN8 = "" Then FN = FN & " | " & Left(TxTBN8, Len(TxTBN8) - 4)
If Not TxTBN9 = "" Then FN = FN & " | " & Left(TxTBN9, Len(TxTBN9) - 4)
If Not TxTBN10 = "" Then FN = FN & " | " & Left(TxTBN10, Len(TxTBN10) - 4)

Bestandsnaam.Hide

CbSendAll = True

Call Factuur_Compleet(PDF, DFA, FN, DSPLEML, CbARC, CbINK, CbAFGH, CbSendAll, CbAB, EM, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN1_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String

Bestandsnaam.Hide

PDF = TxTBN1

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN2_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String

Bestandsnaam.Hide

PDF = TxTBN2

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN3_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String
    
Bestandsnaam.Hide

PDF = TxTBN3

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN4_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String
    
Bestandsnaam.Hide

PDF = TxTBN4

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN5_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String
    
Bestandsnaam.Hide

PDF = TxTBN5

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN6_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String

Bestandsnaam.Hide

PDF = TxTBN6

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN7_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String

Bestandsnaam.Hide

PDF = TxTBN7

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN8_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String
    
Bestandsnaam.Hide

PDF = TxTBN8

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN9_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String
    
Bestandsnaam.Hide

PDF = TxTBN9

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN10_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String
    
Bestandsnaam.Hide

PDF = TxTBN10

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN11_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String
    
Bestandsnaam.Hide

PDF = TxTBN11

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN12_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String
    
Bestandsnaam.Hide

PDF = TxTBN12

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN13_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String
    
Bestandsnaam.Hide

PDF = TxTBN13

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN14_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String
    
Bestandsnaam.Hide

PDF = TxTBN14

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub ObBN15_Click()
    Dim PDF As String
    Dim PDFName(1 To 15) As String
    
Bestandsnaam.Hide

PDF = TxTBN15

Call SearchPDF(PDF, INTERN, WaardenBdNm, WaardenEM, strFolderpath, DFA2, DFA3, DFA4, PDFCount, EM, SN, t, BdNm, SubTxT, DSPLEML, CbAFGH, CbMrg, SRCHPDF, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

End Sub
Private Sub Close1_Click()
Me.Caption = "PDF bestanden uit map verwijderen....."
KillAll
Me.Caption = "Bestandsnaam"
Bestandsnaam.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 1 Then Cancel = 1
End Sub
