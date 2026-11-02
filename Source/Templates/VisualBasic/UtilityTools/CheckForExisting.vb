Imports Inventor

Module CheckForExisting

    Function GetButtonDef(internalName As String) As ButtonDefinition

        ' Check if button exists
        Dim buttonDefinition As Inventor.ButtonDefinition = Nothing
        Try
            buttonDefinition = _inventorApplication.CommandManager.ControlDefinitions.Item(internalName)
        Catch ex As Exception
        End Try

        If Not buttonDefinition Is Nothing Then
            System.Diagnostics.Debug.WriteLine("*******  " & buttonDefinition.InternalName & " button exists")
        End If

        Return buttonDefinition

    End Function

    Function GetRibbonPanel(customRibbonTab As RibbonTab, panelName As String) As RibbonPanel

        Dim panelId As String = "id_" & Replace(panelName, " ", "_")

        ' Check if panel exists
        Dim ribbonPanel As RibbonPanel = Nothing
        Try
            ribbonPanel = customRibbonTab.RibbonPanels.Item(panelId)
        Catch ex As Exception
        End Try

        If Not ribbonPanel Is Nothing Then
            System.Diagnostics.Debug.WriteLine("*******  " & ribbonPanel.InternalName & " panel exists")
        End If

        Return ribbonPanel

    End Function

    Function GetRibbonTab(tabName As String, environmentName As String) As RibbonTab

        Dim internalName As String = "id_" & Replace(tabName, " ", "_")

        ' Check if panel exists
        Dim ribbonTab As RibbonTab = Nothing
        Try
            ribbonTab = _inventorApplication.UserInterfaceManager.Ribbons.Item(environmentName).RibbonTabs.Item(internalName)
        Catch ex As Exception
        End Try

        If Not ribbonTab Is Nothing Then
            System.Diagnostics.Debug.WriteLine("*******  " & tabName & " tab exists")
        End If

        Return ribbonTab

    End Function

End Module
