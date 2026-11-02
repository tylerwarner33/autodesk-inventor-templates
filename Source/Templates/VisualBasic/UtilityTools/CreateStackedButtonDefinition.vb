Imports Inventor
Imports System.Windows.Forms

Module CreateStackedButtonDefinition

    ''' <summary>
    ''' Creates a simple stack of buttons
    ''' See the 'Switch' control available in the Windows panel of the View tab
    ''' </summary> 
    Sub CreatePopUpButtonDefinition(
        environmentName As String,
        customDrawingTab As RibbonTab,
        ribbonPanel As RibbonPanel,
        collection As ObjectCollection,
        defaultButton As ButtonDefinition,
        useLargeIcon As Boolean)

        ribbonPanel = CreateRibbonPanel.GetRibbonPanel(ribbonPanel.DisplayName, customDrawingTab)
        Dim panelName = ribbonPanel.DisplayName
        Dim tabName = customDrawingTab.DisplayName
        Dim defaultButtonName = defaultButton.DisplayName
        Dim collectionCount As Integer = collection.Count

        Try
            ribbonPanel.CommandControls.AddPopup(defaultButton, collection, useLargeIcon, True, "", False)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in CreatePopUp_ButtonDef" & environmentName & ", " &
                tabName & ", " & panelName & ", " & defaultButtonName)
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' Creates a stack of buttons that contains checkboxes that can be toggled on / off
    ''' See the 'Object Visibility' control available in the Visibility panel of the Views tab
    ''' </summary>
    Sub CreateTogglePopUpButtonDefinition(
        environmentName As String,
        customDrawingTab As RibbonTab,
        ribbonPanel As RibbonPanel,
        collection As ObjectCollection,
        defaultButton As ButtonDefinition,
        useLargeIcon As Boolean)

        ribbonPanel = CreateRibbonPanel.GetRibbonPanel(ribbonPanel.DisplayName, customDrawingTab)
        Dim panelName = ribbonPanel.DisplayName
        Dim tabName = customDrawingTab.DisplayName
        Dim defaultButtonName = defaultButton.DisplayName
        Dim collectionCount As Integer = collection.Count

        Try
            ribbonPanel.CommandControls.AddTogglePopup(defaultButton, collection, useLargeIcon, True, "", False)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in CreateToggle_PopUp_ButtonDef" &
                environmentName & ", " & tabName & ", " & panelName & ", " & defaultButtonName)
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' Creates a stack of buttons that changes the button icon as a tool is selected
    ''' See 'Display Mode' drop down (Shaded, Hidden Edge, Wireframe) available on the 'View' tab of the ribbon in parts and assemblies
    ''' </summary>
    Sub CreateButtonPopUpButtonDefinition(
        environmentName As String,
        customDrawingTab As RibbonTab,
        ribbonPanel As RibbonPanel,
        collection As ObjectCollection,
        useLargeIcon As Boolean)

        ribbonPanel = CreateRibbonPanel.GetRibbonPanel(ribbonPanel.DisplayName, customDrawingTab)
        Dim panelName = ribbonPanel.DisplayName
        Dim tabName = customDrawingTab.DisplayName
        Dim collectionCount As Integer = collection.Count

        Try
            ribbonPanel.CommandControls.AddButtonPopup(collection, useLargeIcon, True, "", False)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in CreateButton_PopUp_ButtonDef" &
                environmentName & ", " & tabName & ", " & panelName)
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' Creates a stack of buttons that changes the button to the most recently used button 
    ''' See the 'Circle' drop down in the Sketch tab of the ribbon.
    ''' </summary> 
    Sub CreateSplitMostRecentlyUsedButtonDefinition(
        environmentName As String,
        customDrawingTab As RibbonTab,
        ribbonPanel As RibbonPanel,
        collection As ObjectCollection,
        useLargeIcon As Boolean)

        ribbonPanel = CreateRibbonPanel.GetRibbonPanel(ribbonPanel.DisplayName, customDrawingTab)
        Dim panelName = ribbonPanel.DisplayName
        Dim tabName = customDrawingTab.DisplayName
        Dim collectionCount As Integer = collection.Count

        Try
            ribbonPanel.CommandControls.AddSplitButtonMRU(collection, useLargeIcon, True, "", False)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in CreateSplit_MostRecentlyUsed_ButtonDef" &
                environmentName & ", " & tabName & ", " & panelName)
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

    End Sub

End Module
