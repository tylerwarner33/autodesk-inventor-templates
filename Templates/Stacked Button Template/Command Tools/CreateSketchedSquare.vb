Imports Inventor
Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms

Module CreateSketchedSquare

    Function CreateButton(
        environment As String,
        customDrawingTab As RibbonTab,
        ribbonPanel As RibbonPanel,
        useLargeIcon As Boolean,
        isInButtonStack As Boolean) As ButtonDefinition

        ' Get the images to use for the button0
        Dim largeIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Two_32)
        Dim standardIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Two_16)
        Dim toolTipImage As IPictureDisp = Nothing

        ' This is the text the user sees on the button.
        Dim buttonLabel As String = "New" & vbLf & "Square Symbol"

        ' Text that displays when the user hovers over the button.
        Dim toolTipSimple As String = "Create a new sketched symbol"
        Dim toolTipExpanded As String = Nothing

#Region "Progressive ToolTip"

        ' Set to true to use a progressive tool tip, and false to a simple tool tip.
        Dim useProgressToolTip As Boolean = False

        ' Only used if 'useProgressToolTip = true'.
        If useProgressToolTip = True Then
            toolTipImage = PictureDispConverter.ToIPictureDisp(My.Resources.AU_ToolTip)

            toolTipExpanded = ChrW(&H2022) & " Line1" & vbLf &
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

    ' This is the code that does the real work when your command is executed.
    Sub RunCommandCode()

        Dim drawingDocument As DrawingDocument = _inventorApplication.ActiveDocument
        Dim transientGeometry As TransientGeometry = _inventorApplication.TransientGeometry
        Dim symbolName As String = "Square Symbol"
        Dim sketchedSymbolDefinition As SketchedSymbolDefinition

        Try ' Try to get it.
            sketchedSymbolDefinition = drawingDocument.SketchedSymbolDefinitions.Item(symbolName)
        Catch ' Add it if not found.
            drawingDocument.SketchedSymbolDefinitions.Add(symbolName)
            sketchedSymbolDefinition = drawingDocument.SketchedSymbolDefinitions.Item(symbolName)
        End Try

        Dim drawingSketch As DrawingSketch = Nothing
        ' Open the sketch in edit mode.
        sketchedSymbolDefinition.Edit(drawingSketch)
        drawingSketch.SketchLines.AddAsTwoPointRectangle(transientGeometry.CreatePoint2d(3, 3), transientGeometry.CreatePoint2d(7, 7))

        ' Quit the Edit mode.
        sketchedSymbolDefinition.ExitEdit(True)

        'Insert the sketched symbol on the active sheet.
        drawingDocument.ActiveSheet.SketchedSymbols.Add(symbolName, transientGeometry.CreatePoint2d(5, 10))

    End Sub

End Module
