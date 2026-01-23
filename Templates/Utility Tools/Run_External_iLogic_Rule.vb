Imports System.IO
Imports System.Reflection
Imports Inventor

''' <summary>
''' Utility module for running external iLogic rules from the add-in.
''' Includes automatic registration of the add-in's External Rules folder with iLogic.
''' </summary>
Module Run_External_iLogic_Rule

    Private _externalRulesPathRegistered As Boolean = False
    Private ReadOnly _iLogicAddInGuid As String = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"

    ''' <summary>
    ''' Runs an external iLogic rule by name.
    ''' </summary>
    Public Sub RunExternalRule(ByVal ExternalRuleName As String)

        Try

            ' The application object.
            Dim addIns As ApplicationAddIns = g_inventorApplication.ApplicationAddIns()

            ' Unique ID code for iLogic Addin.
            Dim iLogicAddIn As ApplicationAddIn = addIns.ItemById(_iLogicAddInGuid)

            ' Starts the process.
            iLogicAddIn.Activate()

            ' Executes the rule.
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

    ''' <summary>
    ''' Registers the add-in's External Rules folder with iLogic's External Rule Directories.
    ''' This is called automatically before running any external rule, and only runs once per session.
    ''' </summary>
    Public Sub EnsureExternalRulesPathRegistered()

        ' Only register once per session.
        If _externalRulesPathRegistered Then Return

        Try
            ' Get the path to our External Rules folder (next to the add-in DLL).
            Dim addinAssemblyPath As String = Assembly.GetExecutingAssembly().Location
            Dim addinDirectory As String = System.IO.Path.GetDirectoryName(addinAssemblyPath)
            Dim externalRulesPath As String = System.IO.Path.Combine(addinDirectory, "External Rules")

            ' Check if the External Rules folder exists.
            If Not Directory.Exists(externalRulesPath) Then
                ' No external rules folder deployed, nothing to register.
                _externalRulesPathRegistered = True
                Return
            End If

            ' Get the iLogic add-in
            Dim addIns As ApplicationAddIns = g_inventorApplication.ApplicationAddIns()
            Dim iLogicAddIn As ApplicationAddIn = Nothing

            Try
                iLogicAddIn = addIns.ItemById(_iLogicAddInGuid)
            Catch
                ' iLogic add-in not available.
                _externalRulesPathRegistered = True
                Return
            End Try

            If iLogicAddIn Is Nothing Then
                _externalRulesPathRegistered = True
                Return
            End If

            ' Activate iLogic to access its automation.
            iLogicAddIn.Activate()

            Dim iLogicAuto As Object = iLogicAddIn.Automation
            If iLogicAuto Is Nothing Then
                _externalRulesPathRegistered = True
                Return
            End If

            ' Get the current external rule directories.
            Dim fileOptions As Object = iLogicAuto.FileOptions
            Dim existingPaths As String() = DirectCast(fileOptions.ExternalRuleDirectories, String())

            ' Check if our path is already registered (case-insensitive comparison).
            For Each existingPath As String In existingPaths
                If String.Equals(existingPath, externalRulesPath, StringComparison.OrdinalIgnoreCase) Then
                    ' Already registered.
                    _externalRulesPathRegistered = True
                    Return
                End If
            Next

            ' Add path to the list.
            Dim newPaths As New List(Of String)(existingPaths)
            newPaths.Add(externalRulesPath)
            fileOptions.ExternalRuleDirectories = newPaths.ToArray()

            _externalRulesPathRegistered = True

        Catch ex As Exception
            ' Log but don't fail, the user can still manually configure if needed.
            System.Diagnostics.Debug.WriteLine("Could not auto-register external rules path: " & ex.Message)
            _externalRulesPathRegistered = True
        End Try

    End Sub

    ''' <summary>
    ''' Optionally call this during add-in activation to register the path early.
    ''' This ensures the path is registered even before the user clicks any external rule buttons.
    ''' </summary>
    Public Sub RegisterExternalRulesPathOnActivation()
        EnsureExternalRulesPathRegistered()
    End Sub

End Module
