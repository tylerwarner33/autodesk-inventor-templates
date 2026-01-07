Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms
Imports Inventor
''' <summary>
''' <para>
''' Creates a Button Definition
''' </para>
''' <para>
''' The function within this module defines the label, icons, tool tip, etc. for this button
''' </para>
''' <para>
''' The sub within this module runs a utility to call an external rule when this button is clicked
''' </para>
''' </summary>
''' 
Module Detect_Dark_Light_Theme

    Function CreateButton(environment As String, CustomDrawingTab As RibbonTab, ribbonPanel As RibbonPanel,
                          useLargeIcon As Boolean, isInButtonStack As Boolean) As ButtonDefinition

        'get the images to use for the button
        Dim largeIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Dog_32)
        Dim standardIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Dog_16)
        Dim toolTipImage As IPictureDisp = Nothing

        'this is the text the user sees on the button
        Dim buttonLabel As String = "Detect" & vbLf & "Theme"

        'text that displays when the user hovers over the button
        Dim toolTip_Simple As String = "Runs an external ilogic rule to detect theme color"
        Dim toolTip_Expanded As String = Nothing


#Region "Progressive ToolTip"

        'set to true to use a progressive tool tip, and false to a simple tool tip
        Dim useProgressToolTip As Boolean = False

        'only used if useProgressToolTip = true
        If useProgressToolTip = True Then
            toolTipImage = PictureDispConverter.ToIPictureDisp(My.Resources.Dog_32)
            toolTip_Expanded = Chr(149) & " This tool pops up the application balloon with a message" & vbLf &
                                        Chr(149) & " Line2" & vbLf &
                                        Chr(149) & " Line3" & vbLf &
                                        Chr(149) & " Line4"
        End If

#End Region

        Dim buttonDef As ButtonDefinition
        buttonDef = CreateButtonDefintion.CreateButtonDef(environment, CustomDrawingTab, ribbonPanel, useLargeIcon,
                                                            isInButtonStack, useProgressToolTip,
                                                            buttonLabel, toolTip_Simple, toolTip_Expanded,
                                                            standardIcon, largeIcon, toolTipImage)

        Return buttonDef

    End Function


    'This is the code that does the real work when your command is executed.
    Sub Run_ExternalRule()
        Run_External_iLogic_Rule.RunExternalRule("Detect Dark or Light Theme")
    End Sub

End Module
