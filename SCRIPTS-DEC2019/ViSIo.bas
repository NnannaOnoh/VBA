Attribute VB_Name = "ViSIo"
Option Explicit
 Sub IKWILVISIO()
 Dim xlApp As Object

 Dim bXStarted As Boolean
 Dim enviro As String


 Dim obj As Object
 Dim Controle As Integer
 Dim Controle1 As Integer



Shell ("cmd.exe")

    Controle = MsgBox("wIL JE vISIO?", vbYesNo + vbQuestion, "Controleer bestand")

    If Controle = vbNo Then Exit Sub
    
    Controle1 = MsgBox("wEET JE ZEKER dAT JE vISIO WILT?", vbYesNo + vbQuestion, "Start ")

    If Controle1 = vbNo Then Exit Sub

     On Error Resume Next
     Set xlApp = GetObject(, "visio.Application")
     If Err <> 0 Then
         Application.StatusBar = "Please wait while Excel source is opened ... "
         Set xlApp = CreateObject("visio.Application")
         bXStarted = True
     End If
     On Error GoTo 0
     
     MsgBox ("aLSJEbLIEFT, vISIO")

   
End Sub
