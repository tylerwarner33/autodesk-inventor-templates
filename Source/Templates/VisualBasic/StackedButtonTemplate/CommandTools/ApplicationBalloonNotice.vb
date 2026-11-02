Imports Inventor
Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms

Module ApplicationBalloonNotice

    Function CreateButton(
        environment As String,
        customDrawingTab As RibbonTab,
        ribbonPanel As RibbonPanel,
        useLargeIcon As Boolean,
        isInButtonStack As Boolean) As ButtonDefinition

        ' Get the images to use for the button.
        Dim largeIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Mouse_32)
        Dim standardIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Mouse_16)
        Dim toolTipImage As IPictureDisp = Nothing

        ' This is the text the user sees on the button.
        Dim buttonLabel As String = "Balloon" & vbLf & "Notice"

        ' Text that displays when the user hovers over the button.
        Dim toolTipSimple As String = "Runs an external ilogic rule to display a balloon notice"
        Dim toolTipExpanded As String = Nothing

#Region "Progressive ToolTip"

        ' Set to true to use a progressive tool tip, and false to a simple tool tip.
        Dim useProgressToolTip As Boolean = True

        ' Only used if 'useProgressToolTip = true'.
        If useProgressToolTip = True Then
            toolTipImage = PictureDispConverter.ToIPictureDisp(My.Resources.AU_ToolTip)

            toolTipExpanded = ChrW(&H2022) & " This tool pops up the application balloon with a message" & vbLf &
                ChrW(&H2022) & " Line2" & vbLf &
                ChrW(&H2022) & " Line3" & vbLf &
                ChrW(&H2022) & " Line4"
        End If

#End Region

        Dim buttonDefinition As ButtonDefinition = CreateButtonDefintion.CreateButtonDef(
            environment,
            customDrawingTab,
            ribbonPanel,
            useLargeIcon,
            isInButtonStack,
            useProgressToolTip,
            buttonLabel,
            toolTipSimple,
            toolTipExpanded,
            standardIcon,
            largeIcon,
            toolTipImage)

        Return buttonDefinition

    End Function

    'This is the code that does the real work when your command is executed.
    Sub RunExternalRule()
        RunExternaliLogicRule.RunExternalRule("ApplicationBalloonTip")
    End Sub

End Module
