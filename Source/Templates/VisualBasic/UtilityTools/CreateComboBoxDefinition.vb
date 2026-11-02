Imports Inventor

Module CreateComboBoxDefinition

    Function Create(
        ribbonTab As RibbonTab,
        ribbonPanelName As String,
        isInButtonStack As Boolean,
        isInSlideOut As Boolean,
        buttonLabel As String,
        dropDownWidth As Long,
        toolTipSimple As String) As ComboBoxDefinition

        Dim buttonNameNoSpaces = buttonLabel.Replace(System.Environment.NewLine, "_").Replace(vbLf, "_").Replace(vbCr, "_").Replace(" ", "_")
        Dim internalName As String = buttonNameNoSpaces & "_" & AddInClientId & "_Button_InternalName"
        Dim description As String = buttonNameNoSpaces & "_Button"
        Dim controlDefinitions As Inventor.ControlDefinitions = InventorApplication.CommandManager.ControlDefinitions

        Dim buttonDefinition As ButtonDefinition = CheckForExisting.GetButtonDefinition(internalName)

        If Not buttonDefinition Is Nothing Then
            Return buttonDefinition
            Exit Function
        End If

        Dim ribbonPanel = CreateRibbonPanel.GetRibbonPanel(ribbonPanelName, ribbonTab)

        buttonDefinition = controlDefinitions.AddComboBoxDefinition(
            buttonLabel,
            internalName,
            CommandTypesEnum.kEditMaskCmdType,
            dropDownWidth,
            AddInClientId)

        buttonDefinition.DescriptionText = description
        buttonDefinition.ToolTipText = toolTipSimple

        If isInButtonStack = False Then
            If isInSlideOut = False Then
                ribbonPanel.CommandControls.AddComboBox(buttonDefinition)
            Else
                ribbonPanel.SlideoutControls.AddComboBox(buttonDefinition)
            End If
        End If

        Return buttonDefinition

    End Function

End Module
