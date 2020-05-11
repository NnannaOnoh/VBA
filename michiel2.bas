Attribute VB_Name = "michiel2"
Sub michiel2()

End Sub

    Dim olApp As New Outlook.Application
    Dim olExp As Outlook.Explorer
    Dim olSel As Outlook.Selection
    Dim olNameSpace As Outlook.NameSpace
    Dim olArchive As Outlook.Folder
    Dim intItem As Integer


    Set olExp = olApp.ActiveExplorer
    Set olSel = olExp.Selection
    Set olNameSpace = olApp.GetNamespace("MAPI")
    Set olArchive = olNameSpace.Folders("Facturen").Folders("Postvak IN").Folders("01-Michiel")

    For intItem = 1 To olSel.Count

        olSel.item(intItem).Move olArchive

    Next intItem
