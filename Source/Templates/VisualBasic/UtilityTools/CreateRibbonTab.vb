Imports Inventor

Module CreateRibbonTab

    Function GetRibbon_Tab(tabName As String, myRibbon As Ribbon, tabNextTo As String) As RibbonTab

        Dim ribbons As Ribbons = _inventorApplication.UserInterfaceManager.Ribbons

        Dim ribbonTab As RibbonTab = CheckForExisting.GetRibbonTab(tabName, myRibbon.InternalName)

        Dim internalName As String = "id_" & Replace(tabName, " ", "_")

        If ribbonTab Is Nothing Then
            ribbonTab = myRibbon.RibbonTabs.Add(tabName, internalName, _addInClientID, tabNextTo, False)
            System.Diagnostics.Debug.WriteLine("*******  " & tabName & " tab created")
        End If

        Return ribbonTab

    End Function

End Module
