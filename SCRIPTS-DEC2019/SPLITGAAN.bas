Attribute VB_Name = "SPLITGAAN"
 Sub splits()

    Dim TxTBN() As String
    Dim f As String
    Dim i As Long
    Dim a() As String
    
   Call SaveAtt(strFolderpath, PDFCount, FNABCMPLT, TxTBN1, TxTBN2, TxTBN3, TxTBN4, TxTBN5, TxTBN6, TxTBN7, TxTBN8, TxTBN9, TxTBN10, TxTBN11, TxTBN12, TxTBN13, TxTBN14, TxTBN15)

     MAILSPLIT.FN = ""
    
 MAILSPLIT.TxTBN1 = ""
 MAILSPLIT.TxTBN2 = ""
 MAILSPLIT.TxTBN3 = ""
 MAILSPLIT.TxTBN4 = ""
 MAILSPLIT.TxTBN5 = ""
 MAILSPLIT.TxTBN6 = ""
 MAILSPLIT.TxTBN7 = ""
 MAILSPLIT.TxTBN8 = ""
 MAILSPLIT.TxTBN9 = ""
MAILSPLIT.TxTBN10 = ""
MAILSPLIT.TxTBN11 = ""
MAILSPLIT.TxTBN12 = ""
MAILSPLIT.TxTBN13 = ""
MAILSPLIT.TxTBN14 = ""
MAILSPLIT.TxTBN15 = ""
MAILSPLIT.TxTBN16 = ""
MAILSPLIT.TxTBN17 = ""
MAILSPLIT.TxTBN18 = ""
MAILSPLIT.TxTBN19 = ""
MAILSPLIT.TxTBN20 = ""
MAILSPLIT.TxTBN21 = ""
MAILSPLIT.TxTBN22 = ""
MAILSPLIT.TxTBN23 = ""
MAILSPLIT.TxTBN24 = ""
MAILSPLIT.TxTBN25 = ""

 MAILSPLIT.FNTxT1 = ""
 MAILSPLIT.FNTxT2 = ""
 MAILSPLIT.FNTxT3 = ""
 MAILSPLIT.FNTxT4 = ""
 MAILSPLIT.FNTxT5 = ""
 MAILSPLIT.FNTxT6 = ""
 MAILSPLIT.FNTxT7 = ""
 MAILSPLIT.FNTxT8 = ""
 MAILSPLIT.FNTxT9 = ""
MAILSPLIT.FNTxT10 = ""
MAILSPLIT.FNTxT11 = ""
MAILSPLIT.FNTxT12 = ""
MAILSPLIT.FNTxT13 = ""
MAILSPLIT.FNTxT14 = ""
MAILSPLIT.FNTxT15 = ""
MAILSPLIT.FNTxT16 = ""
MAILSPLIT.FNTxT17 = ""
MAILSPLIT.FNTxT18 = ""
MAILSPLIT.FNTxT19 = ""
MAILSPLIT.FNTxT20 = ""
MAILSPLIT.FNTxT21 = ""
MAILSPLIT.FNTxT22 = ""
MAILSPLIT.FNTxT23 = ""
MAILSPLIT.FNTxT24 = ""
MAILSPLIT.FNTxT25 = ""

  MAILSPLIT.CbBl1 = False
  MAILSPLIT.CbBl2 = False
  MAILSPLIT.CbBl3 = False
  MAILSPLIT.CbBl4 = False
  MAILSPLIT.CbBl5 = False
  MAILSPLIT.CbBl6 = False
  MAILSPLIT.CbBl7 = False
  MAILSPLIT.CbBl8 = False
  MAILSPLIT.CbBl9 = False
 MAILSPLIT.CbBl10 = False
 MAILSPLIT.CbBl11 = False
 MAILSPLIT.CbBl12 = False
 MAILSPLIT.CbBl13 = False
 MAILSPLIT.CbBl14 = False
 MAILSPLIT.CbBl15 = False
 MAILSPLIT.CbBl16 = False
 MAILSPLIT.CbBl17 = False
 MAILSPLIT.CbBl18 = False
 MAILSPLIT.CbBl19 = False
 MAILSPLIT.CbBl20 = False
 MAILSPLIT.CbBl21 = False
 MAILSPLIT.CbBl22 = False
 MAILSPLIT.CbBl23 = False
 MAILSPLIT.CbBl24 = False
 MAILSPLIT.CbBl25 = False

 MAILSPLIT.CbFCT1 = False
 MAILSPLIT.CbFCT2 = False
 MAILSPLIT.CbFCT3 = False
 MAILSPLIT.CbFCT4 = False
 MAILSPLIT.CbFCT5 = False
 MAILSPLIT.CbFCT6 = False
 MAILSPLIT.CbFCT7 = False
 MAILSPLIT.CbFCT8 = False
 MAILSPLIT.CbFCT9 = False
MAILSPLIT.CbFCT10 = False
MAILSPLIT.CbFCT11 = False
MAILSPLIT.CbFCT12 = False
MAILSPLIT.CbFCT13 = False
MAILSPLIT.CbFCT14 = False
MAILSPLIT.CbFCT15 = False
MAILSPLIT.CbFCT16 = False
MAILSPLIT.CbFCT17 = False
MAILSPLIT.CbFCT18 = False
MAILSPLIT.CbFCT19 = False
MAILSPLIT.CbFCT20 = False
MAILSPLIT.CbFCT21 = False
MAILSPLIT.CbFCT22 = False
MAILSPLIT.CbFCT23 = False
MAILSPLIT.CbFCT24 = False
MAILSPLIT.CbFCT25 = False
    
MAILSPLIT.strFolderpath = "H:\Mijn Documenten\merge\pdf\OLAttachments\"

i = 0
fCount = 0

    ReDim a(1 To 2 ^ 15)
            
    f = Dir(strFolderpath & "*.pdf")
    While Len(f)
            i = i + 1
            a(i) = f
            
            fCount = fCount + 1
                       
        f = Dir()
    Wend
            
          'MsgBox fCount  '<--------------------------------------- VERWEIJEDEDERREEENNNNN!!!! Geen nut meer
                      
    If i Then
        ReDim Preserve a(1 To i)
        
        i = 1
            
      If i = 1 Then
      MAILSPLIT.TxTBN1 = a(i)
      MAILSPLIT.FNTxT1 = Left(a(i), Len(a(i)) - 4)
      If fCount = 1 Then GoTo swip
      i = i + 1
      End If
      If i = 2 Then
      MAILSPLIT.TxTBN2 = a(i)
      MAILSPLIT.FNTxT2 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 2 Then GoTo swip
      If i = 3 Then
      MAILSPLIT.TxTBN3 = a(i)
      MAILSPLIT.FNTxT3 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 3 Then GoTo swip
      If i = 4 Then
      MAILSPLIT.TxTBN4 = a(i)
      MAILSPLIT.FNTxT4 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 4 Then GoTo swip
      If i = 5 Then
      MAILSPLIT.TxTBN5 = a(i)
      MAILSPLIT.FNTxT5 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 5 Then GoTo swip
      If i = 6 Then
      MAILSPLIT.TxTBN6 = a(i)
      MAILSPLIT.FNTxT6 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 6 Then GoTo swip
      If i = 7 Then
      MAILSPLIT.TxTBN7 = a(i)
      MAILSPLIT.FNTxT7 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 7 Then GoTo swip
      If i = 8 Then
      MAILSPLIT.TxTBN8 = a(i)
      MAILSPLIT.FNTxT8 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 8 Then GoTo swip
      If i = 9 Then
      MAILSPLIT.TxTBN9 = a(i)
      MAILSPLIT.FNTxT9 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 9 Then GoTo swip
      If i = 10 Then
      MAILSPLIT.TxTBN10 = a(i)
      MAILSPLIT.FNTxT10 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 10 Then GoTo swip
      If i = 11 Then
      MAILSPLIT.TxTBN11 = a(i)
      MAILSPLIT.FNTxT11 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 11 Then GoTo swip
      If i = 12 Then
      MAILSPLIT.TxTBN12 = a(i)
      MAILSPLIT.FNTxT12 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 12 Then GoTo swip
      If i = 13 Then
      MAILSPLIT.TxTBN13 = a(i)
      MAILSPLIT.FNTxT13 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 13 Then GoTo swip
      If i = 14 Then
      MAILSPLIT.TxTBN14 = a(i)
      MAILSPLIT.FNTxT14 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 14 Then GoTo swip
      If i = 15 Then
      MAILSPLIT.TxTBN15 = a(i)
      MAILSPLIT.FNTxT15 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 15 Then GoTo swip
      If i = 16 Then
      MAILSPLIT.TxTBN16 = a(i)
      MAILSPLIT.FNTxT16 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 16 Then GoTo swip
      If i = 17 Then
      MAILSPLIT.TxTBN17 = a(i)
      MAILSPLIT.FNTxT17 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 17 Then GoTo swip
      If i = 18 Then
      MAILSPLIT.TxTBN18 = a(i)
      MAILSPLIT.FNTxT18 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 18 Then GoTo swip
      If i = 19 Then
      MAILSPLIT.TxTBN19 = a(i)
      MAILSPLIT.FNTxT19 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 19 Then GoTo swip
      If i = 20 Then
      MAILSPLIT.TxTBN20 = a(i)
      MAILSPLIT.FNTxT20 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 20 Then GoTo swip
      If i = 21 Then
      MAILSPLIT.TxTBN21 = a(i)
      MAILSPLIT.FNTxT21 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 21 Then GoTo swip
      If i = 22 Then
      MAILSPLIT.TxTBN22 = a(i)
      MAILSPLIT.FNTxT22 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 22 Then GoTo swip
      If i = 23 Then
      MAILSPLIT.TxTBN23 = a(i)
      MAILSPLIT.FNTxT23 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 23 Then GoTo swip
      If i = 24 Then
      MAILSPLIT.TxTBN24 = a(i)
      MAILSPLIT.FNTxT24 = Left(a(i), Len(a(i)) - 4)
      i = i + 1
      End If
      If fCount = 24 Then GoTo swip
      If i = 25 Then
      MAILSPLIT.TxTBN25 = a(i)
      MAILSPLIT.FNTxT25 = Left(a(i), Len(a(i)) - 4)
      End If
swip:

    End If
    
   
   MAILSPLIT.Show
 
 End Sub
