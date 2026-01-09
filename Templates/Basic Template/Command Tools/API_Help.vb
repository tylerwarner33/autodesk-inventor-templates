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
''' The sub within this module defines the code to run when this button is clicked
''' </para>
''' </summary> 

Module API_Help

    ''' <summary>
    ''' <para>
    ''' Creates a button Button Definition
    ''' </para>
    ''' <para>
    ''' Define the label, icons, tool tip, etc. for this button
    ''' </para>
    ''' </summary> 

    Function CreateButton(environment As String, CustomDrawingTab As RibbonTab, ribbonPanel As RibbonPanel,
                          useLargeIcon As Boolean, isInButtonStack As Boolean) As ButtonDefinition

        'get the images to use for the button
        Dim largeIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Cat_32)
        Dim standardIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Cat_16)
        Dim toolTipImage As IPictureDisp = Nothing

        'this is the text the user sees on the button
        Dim buttonLabel As String = "API " & vbLf & "Help"

        'text that displays when the user hovers over the button
        Dim toolTip_Simple As String = "Opens the API Help"
        Dim toolTip_Expanded As String = Nothing

        'set to true to use a progressive tool tip, and false to a simple tool tip
        Dim useProgressToolTip As Boolean = False

#Region "Progressive ToolTip"

        'only used if useProgressToolTip = true
        If useProgressToolTip = True Then
            toolTipImage = PictureDispConverter.ToIPictureDisp(My.Resources.API_ToolTip)

#If #NETFRAMEWORK Then
         toolTip_Expanded = Chr(149) & " This tool opens the API help *.chm file in a seperate window" & vbLf &
                                        "    Use the API help to find:" & vbLf &
                                        Chr(149) & "       API Object Model Reference Information" & vbLf &
                                        Chr(149) & "       Example Code"
#Else
         toolTip_Expanded = ChrW(149) & " This tool opens the API help *.chm file in a seperate window" & vbLf &
                                        "    Use the API help to find:" & vbLf &
                                        ChrW(149) & "       API Object Model Reference Information" & vbLf &
                                        ChrW(149) & "       Example Code"
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

        Dim Version = g_inventorApplication.SoftwareVersion.DisplayVersion
        Dim Array = Split(Version, ".")
        Version = Array(0)
        Dim Build = g_inventorApplication.SoftwareVersion.Major

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
