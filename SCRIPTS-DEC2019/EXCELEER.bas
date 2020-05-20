Attribute VB_Name = "EXCELEER"
Sub Afgehandeld()

    Dim olapp As New Outlook.Application
    Dim olExp As Outlook.Explorer
    Dim olSel As Outlook.Selection
    Dim olNameSpace As Outlook.NameSpace
    Dim olArchive As Outlook.Folder
    Dim intItem As Integer
  
    Set olExp = olapp.ActiveExplorer
    Set olSel = olExp.Selection
    Set olNameSpace = olapp.GetNamespace("MAPI")
    Set olArchive = olNameSpace.Folders("Facturen").Folders("Postvak IN").Folders("Afgehandeld " & (Format(Now, "dd-mm-yyyy")))

    For intItem = 1 To olSel.Count

        olSel.Item(intItem).Move olArchive

    Next intItem
   
    Set olExp = Nothing
    Set olSel = Nothing
    Set olNameSpace = Nothing
    Set olArchive = Nothing
        
End Sub
Sub Retour()

    Dim olapp As New Outlook.Application
    Dim olExp As Outlook.Explorer
    Dim olSel As Outlook.Selection
    Dim olNameSpace As Outlook.NameSpace
    Dim olArchive As Outlook.Folder
    Dim intItem As Integer


    Set olExp = olapp.ActiveExplorer
    Set olSel = olExp.Selection
    Set olNameSpace = olapp.GetNamespace("MAPI")
    Set olArchive = olNameSpace.Folders("Facturen").Folders("Postvak IN").Folders("Afgehandeld " & (Format(Now, "dd-mm-yyyy"))).Folders("Retour leverancier")

    For intItem = 1 To olSel.Count

        olSel.Item(intItem).Move olArchive

    Next intItem
    
End Sub
Sub verwijder()

Call Afgehandeld

Call KNOP3

End Sub
