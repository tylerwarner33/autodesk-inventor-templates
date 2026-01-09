Imports Inventor

Module CheckForExisting

    Function GetButtonDef(InternalName As String) As ButtonDefinition

        ' Check if button exists
        Dim ButtonDef As Inventor.ButtonDefinition = Nothing
        Try
            ButtonDef = g_inventorApplication.CommandManager.ControlDefinitions.Item(InternalName)
        Catch ex As Exception
        End Try

        If Not ButtonDef Is Nothing Then
            System.Diagnostics.Debug.WriteLine("*******  " & ButtonDef.InternalName & " button exists")
        End If

        Return ButtonDef
    End Function

    Function GetRibbonPanel(Custom_RibbonTab As RibbonTab, panelName As String) As RibbonPanel


        Dim panel_ID As String = "id_" & Replace(panelName, " ", "_")

        ' Check if panel exists
        Dim ribbon_Panel As RibbonPanel = Nothing
        Try
            ribbon_Panel = Custom_RibbonTab.RibbonPanels.Item(panel_ID)
        Catch ex As Exception
        End Try

        If Not ribbon_Panel Is Nothing Then
            System.Diagnostics.Debug.WriteLine("*******  " & ribbon_Panel.InternalName & " panel exists")
        End If

        Return ribbon_Panel
    End Function


    Function GetRibbonTab(TabName As String, EnvironmentName As String) As RibbonTab

        Dim internalName As String = "id_" & Replace(TabName, " ", "_")

        ' Check if panel exists
        Dim Ribbon_Tab As RibbonTab = Nothing
        Try
            Ribbon_Tab = g_inventorApplication.UserInterfaceManager.Ribbons.Item(EnvironmentName).RibbonTabs.Item(internalName)
        Catch ex As Exception
        End Try

        If Not Ribbon_Tab Is Nothing Then
            System.Diagnostics.Debug.WriteLine("*******  " & TabName & " tab exists")
        End If

        Return Ribbon_Tab
    End Function

End Module
