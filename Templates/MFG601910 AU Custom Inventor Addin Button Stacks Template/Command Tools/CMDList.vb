Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms
Imports System.Reflection
Imports Inventor

Module CMDList

    Function CreateButton(environment As String, CustomDrawingTab As RibbonTab, ribbonPanel As RibbonPanel,
                          useLargeIcon As Boolean, isInButtonStack As Boolean) As ButtonDefinition

        'get the images to use for the button
        Dim largeIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Snake_32)
        Dim standardIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Snake_16)
        Dim toolTipImage As IPictureDisp = Nothing

        'this is the text the user sees on the button
        Dim buttonLabel As String = "Command" & vbLf & "List"

        'text that displays when the user hovers over the button
        Dim toolTip_Simple As String = "Writes out all Inventor Commands to a text file"
        Dim toolTip_Expanded As String = Nothing


#Region "Progressive ToolTip"

        'set to true to use a progressive tool tip, and false to a simple tool tip
        Dim useProgressToolTip As Boolean = True

        'only used if useProgressToolTip = true
        If useProgressToolTip = True Then
            toolTipImage = PictureDispConverter.ToIPictureDisp(My.Resources.AU_ToolTip)

            toolTip_Expanded = Chr(149) & "  This tool creates a text file and writes all the command information out" & vbLf &
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
    Sub RunCommandCode()

        Dim oFolder As String = "C:\temp\"
        Dim sPath As String = oFolder & "Inventor Command List.txt"

        Dim oCommandMgr As CommandManager = g_inventorApplication.CommandManager
        Dim oControlDefs As ControlDefinitions = oCommandMgr.ControlDefinitions
        Dim oControlDef As ControlDefinition
        Dim objWriter1 As Object

        'create folder
        If Not System.IO.Directory.Exists(oFolder) Then
            System.IO.Directory.CreateDirectory(oFolder)
        End If

        'create file
        If System.IO.File.Exists(sPath) = False Then
            objWriter1 = System.IO.File.CreateText(sPath)
            objWriter1.Close()
        End If

        'edit file
        Dim objWriter As New System.IO.StreamWriter(sPath)
        objWriter.WriteLine("Inventor Command List")
        objWriter.WriteLine("")
        objWriter.WriteLine("Use Example:")
        objWriter.WriteLine("1) Find the command in the list such as:")
        objWriter.WriteLine("             AppViewCubeHomeCmd ")
        objWriter.WriteLine("2) Then add it to the command manager line such as this, and use this line in your code:")
        objWriter.WriteLine("            ThisApplication.CommandManager.ControlDefinitions.Item(" & Chr(34) & "AppViewCubeHomeCmd" & Chr(34) & ").Execute")
        objWriter.WriteLine("____________________________________________________________________________________")
        objWriter.WriteLine("")

        For Each oControlDef In oControlDefs
            If oControlDef.DescriptionText = "" Then
                objWriter.WriteLine(oControlDef.InternalName)
            Else
                objWriter.WriteLine(oControlDef.InternalName & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab &
    "  " & Chr(34) & oControlDef.DescriptionText & Chr(34))
            End If
        Next
        objWriter.Close()
        Process.Start(sPath)




    End Sub
End Module
