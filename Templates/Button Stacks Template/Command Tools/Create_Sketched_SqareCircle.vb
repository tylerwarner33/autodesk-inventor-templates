Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms
Imports Inventor

Module Create_Sketched_SqareCircle

    Function CreateButton(environment As String, CustomDrawingTab As RibbonTab, ribbonPanel As RibbonPanel,
                          useLargeIcon As Boolean, isInButtonStack As Boolean) As ButtonDefinition

        'get the images to use for the button
        Dim largeIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Three_32)
        Dim standardIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Three_16)
        Dim toolTipImage As IPictureDisp = Nothing

        'this is the text the user sees on the button
        Dim buttonLabel As String = "New" & vbLf & "Square Circle Symbol"

        'text that displays when the user hovers over the button
        Dim toolTip_Simple As String = "Create a new sketched symbol"
        Dim toolTip_Expanded As String = Nothing


#Region "Progressive ToolTip"

        'set to true to use a progressive tool tip, and false to a simple tool tip
        Dim useProgressToolTip As Boolean = False

        'only used if useProgressToolTip = true
        If useProgressToolTip = True Then
            toolTipImage = PictureDispConverter.ToIPictureDisp(My.Resources.AU_ToolTip)

#If #NETFRAMEWORK Then
         toolTip_Expanded =
         149) & " Line1" & vbLf &
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
        Dim oDwg As DrawingDocument = g_inventorApplication.ActiveDocument

        Dim oTG As TransientGeometry = g_inventorApplication.TransientGeometry
        Dim oSymbolName As String = "Square Circle Symbol"
        Dim oSSD As SketchedSymbolDefinition

        Try 'try to get it 
            oSSD = oDwg.SketchedSymbolDefinitions.Item(oSymbolName)
        Catch   'add it if not found
            oDwg.SketchedSymbolDefinitions.Add(oSymbolName)
            oSSD = oDwg.SketchedSymbolDefinitions.Item(oSymbolName)
        End Try

        Dim oSketch As DrawingSketch = Nothing
        'Open the sketch in edit mode.
        oSSD.Edit(oSketch)
        oSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(5, 5), 2)
        oSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(5, 5), 1)
        oSketch.SketchLines.AddAsTwoPointRectangle(oTG.CreatePoint2d(3, 3), oTG.CreatePoint2d(7, 7))

        'Quit the Edit mode.
        oSSD.ExitEdit(True)

        'Insert the sketched symbol on the active sheet.
        oDwg.ActiveSheet.SketchedSymbols.Add(oSymbolName, oTG.CreatePoint2d(5, 15))
    End Sub

End Module
