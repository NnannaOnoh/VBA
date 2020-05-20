Attribute VB_Name = "EXCELKNOP"
Sub KNOP1()
If Err Then
MsgBox ("Extra > verwijzingen > Excel")
End If
On Error Resume Next

Dim eApp As Excel.Application
    Set eApp = GetObject(, "Excel.Application")
    eApp.Run "Knop1"
Set eApp = Nothing
End Sub
Sub KNOP2()

On Error Resume Next

Dim eApp As Excel.Application
    Set eApp = GetObject(, "Excel.Application")
    eApp.Run "Knop2"
Set eApp = Nothing
End Sub
Sub KNOP3()

On Error Resume Next

Dim eApp As Excel.Application

        
    Set eApp = GetObject(, "Excel.Application")
    eApp.Run "Knop3"
Set eApp = Nothing
End Sub
Sub KNOP4()

On Error Resume Next

Dim eApp As Excel.Application
    Set eApp = GetObject(, "Excel.Application")
On Error Resume Next
    eApp.Run "Knop4"
Set eApp = Nothing
End Sub
Sub KNOP5()

On Error Resume Next

Dim eApp As Excel.Application
    Set eApp = GetObject(, "Excel.Application")
On Error Resume Next
    eApp.Run "Knop5"
Set eApp = Nothing
End Sub
Sub KNOP7()

On Error Resume Next

Dim eApp As Excel.Application
    Set eApp = GetObject(, "Excel.Application")
On Error Resume Next
    eApp.Run "Knop7"
Set eApp = Nothing
End Sub
Sub circles()

On Error Resume Next

Dim eApp As Excel.Application
    Set eApp = GetObject(, "Excel.Application")
On Error Resume Next
    eApp.Run "circles"
Set eApp = Nothing
End Sub
