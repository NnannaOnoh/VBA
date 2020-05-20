Attribute VB_Name = "Module1"
Sub finden4()

Dim EXCEPT() As String, a As Integer

EM = "no.replynoreply@ziggo.nl"

Exceptions = "no-Reply,noreply,nO.reply,"

            EXCEPT = Split(Exceptions, ",")
            For i = LBound(EXCEPT) To UBound(EXCEPT)
            
    NOREPLY = InStr(1, EM, EXCEPT(i), vbTextCompare)
       
    If NOREPLY > 0 Then
    'CbEM.Value = True
    EM = InputBox("NOREPLY E-MAILADRES", "Geef E-mailadres aan", EM)
    End If

            Next i
MsgBox EM
        
End Sub

      Sub vinden5()

      Dim EXCEPT() As String, a As Integer

      EM = "nnannaonoh@hotmail.com"

      Exceptions = "no-reply,noreply,nO.reply,"

            EXCEPT = Split(Exceptions, ",")
            For i = LBound(EXCEPT) To UBound(EXCEPT)
            
    NOREPLY = InStr(1, EM, EXCEPT(i), vbTextCompare)
       
    If NOREPLY > 0 Then
    'CbEM.Value = True '~food~
    EM = InputBox("NOREPLY E-MAILADRES", "Geef E-mailadres aan", EM)
    'else
    'CbEM.Value = False ~not food~
    End If

            Next i

    MsgBox EM
        
    End Sub
Sub testjes()

        Const Exceptions = "|delivery.moneybird|factuursturen|order2cash|outlook|gmail|ziggo|"
 
 Dim BdNm As String

BdNm = "delivery.moneybird"

nietaangepast:

        If InStr(1, Exceptions, "|" & BdNm & "|", vbTextCompare) Then
  BdNm = InputBox("NOREPLY E-MAILADRES", "Geef E-mailadres aan", BdNm)
Else
  MsgBox "not food"
End If

If InStr(1, Exceptions, "|" & BdNm & "|", vbTextCompare) Then GoTo nietaangepast

End Sub
Sub testjes2222()

Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

pthBdNM = "G:\FIN\11DebCred\Crediteuren\30. Communicatie\301. Emailscripts\Mailbox facturen\Exceptions\FINANCEPF.txt"
strBdNM = FSO.OpenTextFile(pthBdNM).ReadAll

        Exceptions = strBdNM
 
        BdNm = "factuursuren"              '<------------------------- BdNM is al ingevuld vanuit AAPJEPUNTJE

nietaangepast:

If InStr(1, Exceptions, "|" & BdNm & "|", vbTextCompare) Then
        BdNm = InputBox("NOREPLY E-MAILADRES", "Geef E-mailadres aan", BdNm)
        
        If InStr(1, Exceptions, "|" & BdNm & "|", vbTextCompare) Then GoTo nietaangepast
        
Else
        MsgBox "doorgaan met module"
End If


End Sub


Sub usetooo()

Dim tottoo(1 To 15) As String

tottoo(1) = 1
tottoo(2) = 12
tottoo(3) = 13
tottoo(4) = 14
tottoo(5) = 15
tottoo(6) = 16
tottoo(7) = 17


testusetooo (tottoo)


End Sub

Sub testusetooo(tottoo)

MsgBox tottoo(6)






End Sub
