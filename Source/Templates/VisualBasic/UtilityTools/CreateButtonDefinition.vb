Imports Inventor

Module CreateButtonDefinition

    Private _userInputEvents As UserInputEvents

    ''' <summary>
    ''' Creates a simple button.
    ''' </summary>
    Function Create(
        environment As String,
        customTab As RibbonTab,
        ribbonPanelName As RibbonPanel,
        useLargeIcon As Boolean,
        isInButtonStack As Boolean,
        useProgressToolTip As Boolean,
        buttonLabel As String,
        toolTipSimple As String,
        toolTipExpanded As String,
        standardIcon As stdole.IPictureDisp,
        largeIcon As stdole.IPictureDisp,
        toolTipImage As stdole.IPictureDisp) As ButtonDefinition

        Dim buttonNameNoSpaces = buttonLabel.Replace(System.Environment.NewLine, "_").Replace(vbLf, "_").Replace(vbCr, "_").Replace(" ", "_")
        Dim internalName As String = buttonNameNoSpaces & "_" & AddInClientId & "_Button_InternalName"
        Dim description As String = buttonNameNoSpaces & " Button"
        Dim controlDefinitions As Inventor.ControlDefinitions = InventorApplication.CommandManager.ControlDefinitions

        'if using large icon wrap the button name to 2 rows.
        If useLargeIcon = False Then buttonLabel = buttonLabel.Replace(System.Environment.NewLine, " ").Replace(vbLf, " ").Replace(vbCr, " ")

        Dim buttonDefinition As ButtonDefinition = CheckForExisting.GetButtonDefinition(internalName)

        If buttonDefinition Is Nothing Then
            buttonDefinition = controlDefinitions.AddButtonDefinition(buttonLabel, internalName, CommandTypesEnum.kShapeEditCmdType, AddInClientId)
            System.Diagnostics.Debug.WriteLine("*******  " & buttonDefinition.InternalName & " button created")
        End If

        ' Icons for the button.
        buttonDefinition.LargeIcon = largeIcon
        buttonDefinition.StandardIcon = standardIcon

        If useProgressToolTip = True Then
            With buttonDefinition.ProgressiveToolTip
                .IsProgressive = useProgressToolTip
                .Title = buttonLabel
                .Description = description
                .ExpandedDescription = toolTipExpanded
                .Image = toolTipImage
            End With
        End If

        Dim ribbonPanel = CreateRibbonPanel.GetRibbonPanel(ribbonPanelName.DisplayName, customTab)
        ' Button should not be added to panel directly if it is in a button stack.
        If isInButtonStack = False Then
            ribbonPanel.CommandControls.AddButton(buttonDefinition, useLargeIcon, True)
            System.Diagnostics.Debug.WriteLine("*******  " & buttonDefinition.InternalName & " button added to " & ribbonPanel.DisplayName)
        End If

        Return buttonDefinition

    End Function

    Function CreateContext(buttonLabel As String, standardIcon As stdole.IPictureDisp) As ButtonDefinition

        Dim buttonDefinition As ButtonDefinition = Nothing
        Dim controlDefinitions As Inventor.ControlDefinitions = InventorApplication.CommandManager.ControlDefinitions
        Dim addinName = Reflection.Assembly.GetExecutingAssembly.GetName.Name.ToString
        Dim internalName As String = buttonLabel.Replace(System.Environment.NewLine, "_").Replace(vbLf, "_").Replace(vbCr, "_").Replace(" ", "_") & "_" & addinName & "_Button_InternalName"

        Try
            buttonDefinition = controlDefinitions.AddButtonDefinition(buttonLabel, internalName, CommandTypesEnum.kEditMaskCmdType, AddInClientId,,, standardIcon)
        Catch

        End Try

        Return buttonDefinition

    End Function

End Module
