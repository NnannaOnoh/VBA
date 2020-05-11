Attribute VB_Name = "ROUTMAIL"
Sub ROUTEMAIL(ROUTECODE)

strPath = "G:\FIN\11DebCred\Crediteuren\60. Team Input\TIJDELIJK\Contactlijst per route.xlsx"
     On Error Resume Next
     Set xlApp = GetObject(, "Excel.Application")
     If Err <> 0 Then
         Application.StatusBar = "Please wait while Excel source is opened ... "
         Set xlApp = CreateObject("Excel.Application")
         bXStarted = True
     End If
     On Error GoTo 0

     Set xlWb = xlApp.Workbooks.Open(strPath)
     Set xlSheet = xlWb.Sheets(1)

            Cells.Find(What:=ROUTECODE).Activate

            Selection.Offset(0, 1).Select

            BHPAVMAIL = Selection.Value
            
            xlWb.Close False
    
End Sub
