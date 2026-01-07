Imports Inventor
Module CreateButtonDefintion

    Private _UserInputEvents As UserInputEvents


    ''' <summary>
    ''' Creates a simple button
    ''' </summary> 

    Function CreateButtonDef(Environment As String, CustomTab As RibbonTab, ribbonPanelName As RibbonPanel, useLargeIcon As Boolean,
                            isInButtonStack As Boolean, useProgressToolTip As Boolean,
                            buttonLabel As String, toolTip_Simple As String, toolTip_Exapanded As String,
                            standardIcon As stdole.IPictureDisp,
                            largeIcon As stdole.IPictureDisp,
                            toolTipImage As stdole.IPictureDisp) As ButtonDefinition

        Dim ButtonNameNoSpaces = buttonLabel.Replace(vbNewLine, "_").Replace(vbLf, "_").Replace(vbCr, "_").Replace(" ", "_")
        Dim internalName As String = ButtonNameNoSpaces & "_" & g_addInClientID & "_Button_InternalName"
        Dim Description As String = ButtonNameNoSpaces & " Button"
        Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions


        'if using large icon wrap the button name to 2 rows
        If useLargeIcon = False Then buttonLabel = buttonLabel.Replace(vbNewLine, " ").Replace(vbLf, " ").Replace(vbCr, " ")

        Dim toolButtonDef As ButtonDefinition
        toolButtonDef = CheckForExisting.GetButtonDef(internalName)


        If toolButtonDef Is Nothing Then
            toolButtonDef = controlDefs.AddButtonDefinition(buttonLabel, internalName,
                        CommandTypesEnum.kShapeEditCmdType, g_addInClientID)
            System.Diagnostics.Debug.WriteLine("*******  " & toolButtonDef.InternalName & " button created")
        End If

        ' Icons for the button
        toolButtonDef.LargeIcon = largeIcon
        toolButtonDef.StandardIcon = standardIcon

        If useProgressToolTip = True Then
            With toolButtonDef.ProgressiveToolTip
                .IsProgressive = useProgressToolTip
                .Title = buttonLabel
                .Description = Description
                .ExpandedDescription = toolTip_Exapanded
                .Image = toolTipImage
            End With
        End If

        Dim ribbon_Panel = CreateRibbonPanel.GetRibbonPanel(ribbonPanelName.DisplayName, CustomTab)
        'button should not be added to panel directly if it is in a button stack
        If isInButtonStack = False Then
            ribbon_Panel.CommandControls.AddButton(toolButtonDef, useLargeIcon, True)
            System.Diagnostics.Debug.WriteLine("*******  " & toolButtonDef.InternalName & " button added to " & ribbon_Panel.DisplayName)
        End If

        Return toolButtonDef

    End Function


    Function CreateContextButtonDef(buttonLabel As String, standardIcon As stdole.IPictureDisp) As ButtonDefinition

        Dim toolButtonDef As ButtonDefinition = Nothing
        Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions
        Dim AddinName = Reflection.Assembly.GetExecutingAssembly.GetName.Name.ToString
        Dim internalName As String = buttonLabel.Replace(vbNewLine, "_").Replace(vbLf, "_").Replace(vbCr, "_").Replace(" ", "_") & "_" & AddinName & "_Button_InternalName"

        Try
            toolButtonDef = controlDefs.AddButtonDefinition(buttonLabel, internalName,
                                CommandTypesEnum.kEditMaskCmdType, g_addInClientID,,, standardIcon)
        Catch

        End Try

        Return toolButtonDef

    End Function
End Module