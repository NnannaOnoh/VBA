Attribute VB_Name = "EXCELKNOP"
Sub KNOP1()
Dim eApp As Excel.Application
    Set eApp = GetObject(, "Excel.Application")
    eApp.Run "Knop1"
    eApp.Run "circles"
End Sub
Sub KNOP2()
Dim eApp As Excel.Application
    Set eApp = GetObject(, "Excel.Application")
    eApp.Run "Knop2"
End Sub
Sub KNOP5()
Dim eApp As Excel.Application
    Set eApp = GetObject(, "Excel.Application")
    eApp.Run "Knop5"
End Sub
Sub circles()
Dim eApp As Excel.Application
    Set eApp = GetObject(, "Excel.Application")
    eApp.Run "circles"
End Sub
