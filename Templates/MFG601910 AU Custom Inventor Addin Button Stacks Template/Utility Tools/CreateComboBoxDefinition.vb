Imports Inventor
Module CreateComboBoxDefintion

    Function CreateButtonDef(Ribbon_Tab As RibbonTab, ribbonPanelName As String,
                            isInButtonStack As Boolean,
                            isInSlideOut As Boolean,
                            buttonLabel As String,
                            dropDownWidth As Long,
                            toolTip_Simple As String) As ComboBoxDefinition

        Dim ButtonNameNoSpaces = buttonLabel.Replace(vbNewLine, "_").Replace(vbLf, "_").Replace(vbCr, "_").Replace(" ", "_")
        Dim internalName As String = ButtonNameNoSpaces & "_" & g_addInClientID & "_Button_InternalName"
        Dim description As String = ButtonNameNoSpaces & "_Button"
        Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions

        Dim toolButtonDef As ButtonDefinition
        toolButtonDef = CheckForExisting.GetButtonDef(internalName)

        If Not toolButtonDef Is Nothing Then
            Return toolButtonDef
            Exit Function
        End If

        Dim ribbon_Panel = CreateRibbonPanel.GetRibbonPanel(ribbonPanelName, Ribbon_Tab)


        toolButtonDef = controlDefs.AddComboBoxDefinition(buttonLabel, internalName,
                                CommandTypesEnum.kEditMaskCmdType, dropDownWidth, g_addInClientID)


        toolButtonDef.DescriptionText = description
        toolButtonDef.ToolTipText = toolTip_Simple

        If isInButtonStack = False Then
            If isInSlideOut = False Then
                ribbon_Panel.CommandControls.AddComboBox(toolButtonDef)
            Else
                ribbon_Panel.SlideoutControls.AddComboBox(toolButtonDef)
            End If
        End If

        Return toolButtonDef

    End Function

End Module