Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms
Imports Inventor

Module Install_Path

    Function CreateButton(environment As String, CustomDrawingTab As RibbonTab, ribbonPanel As RibbonPanel,
                          useLargeIcon As Boolean, isInButtonStack As Boolean) As ButtonDefinition

        'get the images to use for the button
        Dim largeIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Horse_32)
        Dim standardIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Horse_16)
        Dim toolTipImage As IPictureDisp = Nothing

        'this is the text the user sees on the button
        Dim buttonLabel As String = "Check" & vbLf & "Install Path"

        'text that displays when the user hovers over the button
        Dim toolTip_Simple As String = "Displays the install path"
        Dim toolTip_Expanded As String = Nothing


#Region "Progressive ToolTip"

        'set to true to use a progressive tool tip, and false to a simple tool tip
        Dim useProgressToolTip As Boolean = False

        'only used if useProgressToolTip = true
        If useProgressToolTip = True Then
            toolTipImage = PictureDispConverter.ToIPictureDisp(My.Resources.AU_ToolTip)

#If #NETFRAMEWORK Then
         toolTip_Expanded = Chr(149) & " Line1" & vbLf &
                                     Chr(149) & " Line2" & vbLf &
                                     Chr(149) & " Line3" & vbLf &
                                     Chr(149) & " Line4"
#Else
         toolTip_Expanded = ChrW(149) & " Line1" & vbLf &
                                     ChrW(149) & " Line2" & vbLf &
                                     ChrW(149) & " Line3" & vbLf &
                                     ChrW(149) & " Line4"
#End If
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
    Sub RunCommandCode()
        MsgBox(g_inventorApplication.InstallPath)
    End Sub

End Module
