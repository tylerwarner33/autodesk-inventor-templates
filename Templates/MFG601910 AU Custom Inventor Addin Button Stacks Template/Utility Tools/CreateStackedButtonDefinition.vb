Imports System.Windows.Forms
Imports Inventor
Module CreateStackedButtonDefinition

    ''' <summary>
    ''' <para>
    ''' Creates a simple stack of buttons
    ''' </para>
    ''' <para>
    ''' See the 'Switch' control available in the Windows panel of the View tab
    ''' </para>
    ''' </summary> 
    Sub Create_PopUp_ButtonDef(EnvironmentName As String, CustomDrawingTab As RibbonTab, Ribbon_Panel As RibbonPanel,
                             Collection As ObjectCollection, DefaultButton As ButtonDefinition, UseLargeIcon As Boolean)


        Ribbon_Panel = CreateRibbonPanel.GetRibbonPanel(Ribbon_Panel.DisplayName, CustomDrawingTab)
        Dim PanelName = Ribbon_Panel.DisplayName
        Dim TabName = CustomDrawingTab.DisplayName
        Dim DefaultButtonName = DefaultButton.DisplayName
        Dim CollectionCount As Integer = Collection.Count

        Try
            Ribbon_Panel.CommandControls.AddPopup(DefaultButton, Collection, UseLargeIcon, True, "", False)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in CreatePopUp_ButtonDef" & EnvironmentName & ", " &
                                               TabName & ", " & PanelName & ", " & DefaultButtonName)
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try



    End Sub

    ''' <summary>
    ''' <para>
    ''' Creates a stack of buttons that contains checkboxes that can be toggled on / off
    ''' </para>
    ''' <para>
    ''' See the 'Object Visibility' control available in the Visibility panel of the Views tab
    ''' </para>
    ''' </summary> 
    ''' 
    Sub CreateToggle_PopUp_ButtonDef(EnvironmentName As String, CustomDrawingTab As RibbonTab, Ribbon_Panel As RibbonPanel,
                                    Collection As ObjectCollection, DefaultButton As ButtonDefinition, UseLargeIcon As Boolean)


        Ribbon_Panel = CreateRibbonPanel.GetRibbonPanel(Ribbon_Panel.DisplayName, CustomDrawingTab)
        Dim PanelName = Ribbon_Panel.DisplayName
        Dim TabName = CustomDrawingTab.DisplayName
        Dim DefaultButtonName = DefaultButton.DisplayName
        Dim CollectionCount As Integer = Collection.Count

        Try
            Ribbon_Panel.CommandControls.AddTogglePopup(DefaultButton, Collection, UseLargeIcon, True, "", False)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in CreateToggle_PopUp_ButtonDef" &
                                               EnvironmentName & ", " & TabName & ", " & PanelName & ", " & DefaultButtonName)
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' <para>
    ''' Creates a stack of buttons that changes the button icon as a tool is selected
    ''' </para>
    ''' <para>
    ''' See 'Display Mode' drop down (Shaded, Hidden Edge, Wireframe) available on the 'View' tab of the ribbon in parts and assemblies
    ''' </para>
    ''' </summary> 

    Sub Create_Button_PopUp_ButtonDef(EnvironmentName As String, CustomDrawingTab As RibbonTab, Ribbon_Panel As RibbonPanel,
                                     Collection As ObjectCollection, UseLargeIcon As Boolean)

        Ribbon_Panel = CreateRibbonPanel.GetRibbonPanel(Ribbon_Panel.DisplayName, CustomDrawingTab)
        Dim PanelName = Ribbon_Panel.DisplayName
        Dim TabName = CustomDrawingTab.DisplayName
        Dim CollectionCount As Integer = Collection.Count

        Try
            Ribbon_Panel.CommandControls.AddButtonPopup(Collection, UseLargeIcon, True, "", False)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in CreateButton_PopUp_ButtonDef" &
                                               EnvironmentName & ", " & TabName & ", " & PanelName)
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try


    End Sub



    ''' <summary>
    ''' <para>
    ''' Creates a stack of buttons that changes the button to the most recently used button 
    ''' </para>
    ''' <para>
    ''' See the 'Circle' drop down in the Sketch tab of the ribbon.
    ''' </para>
    ''' </summary> 
    Sub CreateSplit_MostRecentlyUsed_ButtonDef(EnvironmentName As String, CustomDrawingTab As RibbonTab, Ribbon_Panel As RibbonPanel,
                                              Collection As ObjectCollection, UseLargeIcon As Boolean)


        Ribbon_Panel = CreateRibbonPanel.GetRibbonPanel(Ribbon_Panel.DisplayName, CustomDrawingTab)
        Dim PanelName = Ribbon_Panel.DisplayName
        Dim TabName = CustomDrawingTab.DisplayName
        Dim CollectionCount As Integer = Collection.Count

        Try
            Ribbon_Panel.CommandControls.AddSplitButtonMRU(Collection, UseLargeIcon, True, "", False)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in CreateSplit_MostRecentlyUsed_ButtonDef" &
                                               EnvironmentName & ", " & TabName & ", " & PanelName)
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

    End Sub



End Module