Imports Inventor

Module CreateButtonDefintion

    Private _userInputEvents As UserInputEvents

    ''' <summary>
    ''' Creates a simple button.
    ''' </summary>
    Function CreateButtonDef(
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
        Dim internalName As String = buttonNameNoSpaces & "_" & _addInClientID & "_Button_InternalName"
        Dim description As String = buttonNameNoSpaces & " Button"
        Dim controlDefs As Inventor.ControlDefinitions = _inventorApplication.CommandManager.ControlDefinitions

        'if using large icon wrap the button name to 2 rows.
        If useLargeIcon = False Then buttonLabel = buttonLabel.Replace(System.Environment.NewLine, " ").Replace(vbLf, " ").Replace(vbCr, " ")

        Dim toolButtonDef As ButtonDefinition = CheckForExisting.GetButtonDef(internalName)

        If toolButtonDef Is Nothing Then
            toolButtonDef = controlDefs.AddButtonDefinition(buttonLabel, internalName, CommandTypesEnum.kShapeEditCmdType, _addInClientID)
            System.Diagnostics.Debug.WriteLine("*******  " & toolButtonDef.InternalName & " button created")
        End If

        ' Icons for the button.
        toolButtonDef.LargeIcon = largeIcon
        toolButtonDef.StandardIcon = standardIcon

        If useProgressToolTip = True Then
            With toolButtonDef.ProgressiveToolTip
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
            ribbonPanel.CommandControls.AddButton(toolButtonDef, useLargeIcon, True)
            System.Diagnostics.Debug.WriteLine("*******  " & toolButtonDef.InternalName & " button added to " & ribbonPanel.DisplayName)
        End If

        Return toolButtonDef

    End Function

    Function CreateContextButtonDef(buttonLabel As String, standardIcon As stdole.IPictureDisp) As ButtonDefinition

        Dim toolButtonDef As ButtonDefinition = Nothing
        Dim controlDefs As Inventor.ControlDefinitions = _inventorApplication.CommandManager.ControlDefinitions
        Dim addinName = Reflection.Assembly.GetExecutingAssembly.GetName.Name.ToString
        Dim internalName As String = buttonLabel.Replace(System.Environment.NewLine, "_").Replace(vbLf, "_").Replace(vbCr, "_").Replace(" ", "_") & "_" & addinName & "_Button_InternalName"

        Try
            toolButtonDef = controlDefs.AddButtonDefinition(buttonLabel, internalName, CommandTypesEnum.kEditMaskCmdType, _addInClientID,,, standardIcon)
        Catch

        End Try

        Return toolButtonDef

    End Function

End Module
