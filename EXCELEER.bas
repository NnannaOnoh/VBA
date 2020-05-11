Attribute VB_Name = "EXCELEER"
Sub Afgehandeld()

    Dim olApp As New Outlook.Application
    Dim olExp As Outlook.Explorer
    Dim olSel As Outlook.Selection
    Dim olNameSpace As Outlook.NameSpace
    Dim olArchive As Outlook.Folder
    Dim intItem As Integer
  
    Set olExp = olApp.ActiveExplorer
    Set olSel = olExp.Selection
    Set olNameSpace = olApp.GetNamespace("MAPI")
    Set olArchive = olNameSpace.Folders("Facturen").Folders("Postvak IN").Folders("Afgehandeld " & (Format(Now, "dd-mm-yyyy")))

    For intItem = 1 To olSel.Count

        olSel.Item(intItem).Move olArchive

    Next intItem

    'Call KNOP1
        
End Sub
Sub Retour()

    Dim olApp As New Outlook.Application
    Dim olExp As Outlook.Explorer
    Dim olSel As Outlook.Selection
    Dim olNameSpace As Outlook.NameSpace
    Dim olArchive As Outlook.Folder
    Dim intItem As Integer


    Set olExp = olApp.ActiveExplorer
    Set olSel = olExp.Selection
    Set olNameSpace = olApp.GetNamespace("MAPI")
    Set olArchive = olNameSpace.Folders("Facturen").Folders("Postvak IN").Folders("Afgehandeld " & (Format(Now, "dd-mm-yyyy"))).Folders("Retour leverancier")

    For intItem = 1 To olSel.Count

        olSel.Item(intItem).Move olArchive

    Next intItem

    'Call KNOP2
    
End Sub
Sub AHMETF()

    Dim olApp As New Outlook.Application
    Dim olExp As Outlook.Explorer
    Dim olSel As Outlook.Selection
    Dim olNameSpace As Outlook.NameSpace
    Dim olArchive As Outlook.Folder
    Dim intItem As Integer


    Set olExp = olApp.ActiveExplorer
    Set olSel = olExp.Selection
    Set olNameSpace = olApp.GetNamespace("MAPI")
    Set olArchive = olNameSpace.Folders("Facturen").Folders("Postvak IN").Folders("002Facturen ouder dan een week")

    For intItem = 1 To olSel.Count

        olSel.Item(intItem).Move olArchive

    Next intItem

    'Call KNOP5
    
End Sub
