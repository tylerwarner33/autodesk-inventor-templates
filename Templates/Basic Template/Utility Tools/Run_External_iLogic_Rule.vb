
Imports Inventor
'Imports Autodesk.iLogic.Interfaces
'Imports Autodesk.iLogic.Automation

Module Run_External_iLogic_Rule


    Public Sub RunExternalRule(ByVal ExternalRuleName As String)

        Try

            ' The application object.
            Dim addIns As ApplicationAddIns = g_inventorApplication.ApplicationAddIns()

            ' Unique ID code for iLogic Addin
            Dim iLogicAddIn As ApplicationAddIn = addIns.ItemById("{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}")

            ' Starts the process
            iLogicAddIn.Activate()

            ' Executes the rule
            iLogicAddIn.Automation.RunExternalRule(g_inventorApplication.ActiveDocument, ExternalRuleName)

        Catch ex As Exception

            MsgBox("Error launching external rule: " & vbLf & "     " & ExternalRuleName _
                   & vbLf & vbLf & "Ensure the iLogic rule exists, and that the" _
                   & vbLf & "configuration includes the path to the rules." _
                   & vbLf & vbLf & "see Tools tab > Options flyout button > iLogic Configuration button" _
                   & vbLf & "    If needed, add an external rules folder path," _
                   & vbLf & "    and ensure that the rule is found in this folder.")
            Return

        End Try

    End Sub

End Module
