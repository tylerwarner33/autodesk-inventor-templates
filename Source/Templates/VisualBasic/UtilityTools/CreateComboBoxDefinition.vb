Imports Inventor

Module CreateComboBoxDefintion

    Function CreateButtonDef(
        ribbonTab As RibbonTab,
        ribbonPanelName As String,
        isInButtonStack As Boolean,
        isInSlideOut As Boolean,
        buttonLabel As String,
        dropDownWidth As Long,
        toolTipSimple As String) As ComboBoxDefinition

        Dim buttonNameNoSpaces = buttonLabel.Replace(System.Environment.NewLine, "_").Replace(vbLf, "_").Replace(vbCr, "_").Replace(" ", "_")
        Dim internalName As String = buttonNameNoSpaces & "_" & _addInClientID & "_Button_InternalName"
        Dim description As String = buttonNameNoSpaces & "_Button"
        Dim controlDefs As Inventor.ControlDefinitions = _inventorApplication.CommandManager.ControlDefinitions

        Dim toolButtonDef As ButtonDefinition = CheckForExisting.GetButtonDef(internalName)

        If Not toolButtonDef Is Nothing Then
            Return toolButtonDef
            Exit Function
        End If

        Dim ribbonPanel = CreateRibbonPanel.GetRibbonPanel(ribbonPanelName, ribbonTab)

        toolButtonDef = controlDefs.AddComboBoxDefinition(
            buttonLabel,
            internalName,
            CommandTypesEnum.kEditMaskCmdType,
            dropDownWidth,
            _addInClientID)

        toolButtonDef.DescriptionText = description
        toolButtonDef.ToolTipText = toolTipSimple

        If isInButtonStack = False Then
            If isInSlideOut = False Then
                ribbonPanel.CommandControls.AddComboBox(toolButtonDef)
            Else
                ribbonPanel.SlideoutControls.AddComboBox(toolButtonDef)
            End If
        End If

        Return toolButtonDef

    End Function

End Module
