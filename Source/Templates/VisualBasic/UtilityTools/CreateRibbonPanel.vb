Imports Inventor
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Module CreateRibbonPanel

    Function GetRibbonPanel(panelName As String, myRibbonTab As RibbonTab) As RibbonPanel

        Dim customRibbonTab As RibbonTab = CheckForExisting.GetRibbonTab(myRibbonTab.DisplayName, myRibbonTab.Parent.InternalName)

        Dim ribbonPanel As RibbonPanel = CheckForExisting.GetRibbonPanel(customRibbonTab, panelName)

        Dim panelId As String = "id_" & Replace(panelName, " ", "_")

        If ribbonPanel Is Nothing Then
            ribbonPanel = customRibbonTab.RibbonPanels.Add(panelName, panelId, _addInClientID)
            System.Diagnostics.Debug.WriteLine("*******  " & panelName & " panel created")
        End If

        Return ribbonPanel

    End Function

End Module
