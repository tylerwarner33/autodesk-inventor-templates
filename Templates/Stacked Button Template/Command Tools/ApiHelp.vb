Imports Inventor

''' <summary>
''' Creates a Button Definition.
''' The function within this module defines the label, icons, tool tip, etc. for this button
''' The sub within this module defines the code to run when this button is clicked
''' </summary> 
Module ApiHelp

    'This function is where the button is defined
    Function CreateButton(
        environment As String,
        CustomDrawingTab As RibbonTab,
        ribbonPanel As RibbonPanel,
        useLargeIcon As Boolean,
        isInButtonStack As Boolean) As ButtonDefinition

        ' Get the images to use for the button.
        Dim largeIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Cat_32)
        Dim standardIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Cat_16)
        Dim toolTipImage As IPictureDisp = Nothing

        ' This is the text the user sees on the button.
        Dim buttonLabel As String = "API " & vbLf & "Help"

        ' Text that displays when the user hovers over the button.
        Dim toolTipSimple As String = "Opens the API Help"
        Dim toolTipExpanded As String = Nothing

#Region "Progressive ToolTip"

        ' Set to true to use a progressive tool tip, and false to a simple tool tip.
        Dim useProgressToolTip As Boolean = True

        ' Only used if 'useProgressToolTip = true'.
        If useProgressToolTip = True Then
            toolTipImage = PictureDispConverter.ToIPictureDisp(My.Resources.API_ToolTip)

            toolTipExpanded = ChrW(&H2022) & " This tool opens the API help *.chm file in a seperate window" & vbLf &
                ChrW(&H2022) & " Line2" & vbLf &
                ChrW(&H2022) & " Line3" & vbLf &
                ChrW(&H2022) & " Line4"
        End If

#End Region

        Dim buttonDefinition As ButtonDefinition = CreateButtonDefintion.CreateButtonDef(
            environment,
            CustomDrawingTab,
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
    Sub RunCommandCode()

        Dim Version = _inventorApplication.SoftwareVersion.DisplayVersion
        Dim Array = Split(Version, ".")
        Version = Array(0)
        Dim Build = _inventorApplication.SoftwareVersion.Major

        Dim Filename = "C:\Users\Public\Documents\Autodesk\Inventor " &
        Version & "\Local Help\ADMAPI_" & Build & "_0.chm"

        If System.IO.File.Exists(Filename) = True Then
#If #NETFRAMEWORK Then
            Process.Start(Filename)
#Else
            Process.Start(New ProcessStartInfo(Filename) With {.UseShellExecute = True})
#End If
        Else
            MsgBox("File Does Not Exist" & vbLf & Filename)
        End If

    End Sub

End Module
