Imports Inventor
Module CreateRibbonTab

    Function GetRibbon_Tab(tabName As String, myRibbon As Ribbon, Tab_NextTo As String) As RibbonTab

        '' Get the correct ribbon
        Dim Ribbons As Ribbons = g_inventorApplication.UserInterfaceManager.Ribbons

        Dim Ribbon_Tab As RibbonTab
        Ribbon_Tab = CheckForExisting.GetRibbonTab(tabName, myRibbon.InternalName)

        Dim internalName As String = "id_" & Replace(tabName, " ", "_")

        If Ribbon_Tab Is Nothing Then
            Ribbon_Tab = myRibbon.RibbonTabs.Add(tabName, internalName, g_addInClientID, Tab_NextTo, False)
            System.Diagnostics.Debug.WriteLine("*******  " & tabName & " tab created")
        End If

        Return Ribbon_Tab
    End Function


End Module