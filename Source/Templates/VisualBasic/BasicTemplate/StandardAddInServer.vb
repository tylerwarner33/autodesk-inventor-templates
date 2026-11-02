Imports Inventor
Imports Microsoft.Win32
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

' Define global variables.
Public Module Globals

    Public _inventorApplication As Inventor.Application
    ''' <summary>
    ''' Unique ID for this add-in.
    ''' </summary>
    Public Const _simpleAddInClientID As String = "CLIENT_GUID_PLACEHOLDER"
    Public Const _addInClientID As String = "{" & _simpleAddInClientID & "}"

End Module

Namespace InventorAddIn

    <ProgIdAttribute("InventorAddIn.StandardAddInServer"), GuidAttribute(_simpleAddInClientID)>
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        Private WithEvents _applicationEvents As ApplicationEvents
        Private WithEvents _userInterfaceEvents As UserInterfaceEvents
        Private WithEvents _userInputEvents As UserInputEvents
        Private WithEvents _transactionEvents As TransactionEvents

        ' Define the command buttons.
        Private WithEvents _apiHelpButton As ButtonDefinition
        Private WithEvents _detectInventorThemeButton As ButtonDefinition

        Private Sub AddToUserInterface()

#Region "Setup Information"

            ' Create a object collection and Control Definition array to be used with button stacks.
            Dim buttonObjectCollection As ObjectCollection = _inventorApplication.TransientObjects.CreateObjectCollection
            Dim controlDefinition(0 To 99) As ControlDefinition

            ' Create list of environments to cycle through.
            Dim environmentList As New List(Of String)({"Drawing", "Assembly", "Part"})

            ' Define custom roibbon tab prefix and suffix.
            ' This will be combined with the EnvironmentList to create ribbon tabs like "ACME Drawing Tools", "ACME Part Tools", etc.
            Dim ribbonTabNamePrefix As String = "ACME"
            Dim ribbonTabNameSufffix As String = "Tools"

            ' Create list of panels to create on the custom ribbon tab.
            Dim customPanelList As New List(Of String)({"General Tools", "External Rule Buttons"})
#End Region

#Region "Create Custom Tabs and Panels for each environment in the list"

            ' Create Custom Tabs and Panels for each environment in the list
            ' Example:
            ''      Environment = "Drawing"
            ''      Prefix = 'ACME'
            ''      Suffix = "Tools"
            ''      will create a custom tab called: ACME Drawing Tools

            Dim allRibbons As Ribbons = _inventorApplication.UserInterfaceManager.Ribbons
            Dim visibleTabsCollection As ObjectCollection = _inventorApplication.TransientObjects.CreateObjectCollection

            Dim customDrawingTab As RibbonTab = Nothing
            Dim customAssemblyTab As RibbonTab = Nothing
            Dim customPartTab As RibbonTab = Nothing

            Dim drawingGeneralToolsPanel As RibbonPanel = Nothing
            Dim drawingExternalRuleButtonsPanel As RibbonPanel = Nothing
            Dim assemblyGeneralToolsPanel As RibbonPanel = Nothing
            Dim assemblyExternalRuleButtonsPanel As RibbonPanel = Nothing
            Dim partGeneralToolsPanel As RibbonPanel = Nothing
            Dim partExternalRuleButtonsPanel As RibbonPanel = Nothing


            ' Iterate through the list.
            For i = 0 To environmentList.Count - 1

                ' Get the ribbon for each environment ( example: Drawing, Assembly, Part).
                Dim myRibbon As Ribbon = allRibbons.Item(environmentList.Item(i))

                ' Make sure the collection is empty.
                visibleTabsCollection.Clear()

                ' Collect only the visible tabs in the ribbon.
                For Each ribbonTab In myRibbon.RibbonTabs
                    If ribbonTab.Visible = False Then Continue For
                    visibleTabsCollection.Add(ribbonTab)
                Next

                ' Get count of tabs in collection.
                Dim iCount As Integer = visibleTabsCollection.Count

                ' Get the last tab in the collection.
                Dim lastTab As String = visibleTabsCollection.Item(iCount).InternalName

                ' Create new tab name.
                Dim newTabName As String = ribbonTabNamePrefix & " " & myRibbon.InternalName & " " & ribbonTabNameSufffix

                ' Create new tab.
                Dim customTab As RibbonTab = CreateRibbonTab.GetRibbon_Tab(newTabName, myRibbon, lastTab)

                ' Create the tab and panels for each environment.
                If i = 0 Then
                    customDrawingTab = customTab
                    drawingGeneralToolsPanel = CreateRibbonPanel.GetRibbonPanel(customPanelList.Item(0), customTab)
                    drawingExternalRuleButtonsPanel = CreateRibbonPanel.GetRibbonPanel(customPanelList.Item(1), customTab)
                ElseIf i = 1 Then
                    customAssemblyTab = customTab
                    assemblyGeneralToolsPanel = CreateRibbonPanel.GetRibbonPanel(customPanelList.Item(0), customTab)
                    assemblyExternalRuleButtonsPanel = CreateRibbonPanel.GetRibbonPanel(customPanelList.Item(1), customTab)
                ElseIf i = 2 Then
                    customPartTab = customTab
                    partGeneralToolsPanel = CreateRibbonPanel.GetRibbonPanel(customPanelList.Item(0), customTab)
                    partExternalRuleButtonsPanel = CreateRibbonPanel.GetRibbonPanel(customPanelList.Item(1), customTab)
                End If

            Next
#End Region

#Region "Create Single Buttons"

            ' Add this button to General Tools tab of 3 different environments.
            _apiHelpButton = ApiHelp.CreateButton("Drawing", customDrawingTab, drawingGeneralToolsPanel, True, False)
            _apiHelpButton = ApiHelp.CreateButton("Assembly", customAssemblyTab, assemblyGeneralToolsPanel, True, False)
            _apiHelpButton = ApiHelp.CreateButton("Part", customPartTab, partGeneralToolsPanel, True, False)

            ' Add button to just one tab.
            _detectInventorThemeButton = DetectInventorTheme.CreateButton("Drawing", customDrawingTab, drawingExternalRuleButtonsPanel, True, False)

#End Region

        End Sub

#Region "Button 'on click' Events"

        ' Add button click events to this region.

        Private Sub ApiHelp_Button_OnExecute(context As NameValueMap) Handles _apiHelpButton.OnExecute
            ApiHelp.RunCommandCode()
        End Sub

        Private Sub DetectInventorTheme_Button_OnExecute(context As NameValueMap) Handles _detectInventorThemeButton.OnExecute
            DetectInventorTheme.RunExternalRule()
        End Sub
#End Region

#Region "Applicaton Events"

        Private Sub Events_OnNewDocument(
            ByVal documentObject As Inventor._Document,
            ByVal beforeOrAfter As Inventor.EventTimingEnum,
            ByVal context As Inventor.NameValueMap,
            ByRef handlingCode As Inventor.HandlingCodeEnum)

            If documentObject Is Nothing Then Exit Sub

            If documentObject.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                If beforeOrAfter = EventTimingEnum.kAfter Then
                    ' Run a command that has no button.
                    ' Uncomment the following line to see the Welcome command module run on this event.
                    'Welcome.AddinHello(documentObject.FullFileName, EventTimingEnum.kAfter.ToString)
                End If
            End If

        End Sub

        Private Sub Events_OnOpenDocument(
            documentObject As _Document,
            fullDocumentName As String,
            beforeOrAfter As EventTimingEnum,
            context As NameValueMap,
            ByRef handlingCode As HandlingCodeEnum)

            If documentObject Is Nothing Then Exit Sub
            If documentObject IsNot _inventorApplication.ActiveDocument Then Exit Sub ' Only trigger on the active file.
            If Not documentObject.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then Exit Sub

            'run this after a document is open
            If beforeOrAfter = EventTimingEnum.kAfter Then
                ' Run a command that has no button.
                ' Uncomment the following line to see the Welcome command module run on this event.
                'Welcome.AddinHello(documentObject.FullFileName, EventTimingEnum.kAfter.ToString)
            End If

        End Sub

        Private Sub Events_OnSaveDocument(
            documentObject As _Document,
            beforeOrAfter As EventTimingEnum,
            context As NameValueMap,
            ByRef handlingCode As HandlingCodeEnum)

            If documentObject.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                If beforeOrAfter = EventTimingEnum.kBefore Then
                    ' Run a command that has no button.
                    ' Uncomment the following line to see the Welcome command module run on this event.
                    'Welcome.AddinHello(documentObject.FullFileName, EventTimingEnum.kAfter.ToString)
                End If
            End If

        End Sub

#End Region

#Region "ApplicationAddInServer Members"

        ''' <summary>
        ''' This method is called by Inventor when it loads the AddIn.
        ''' The AddInSiteObject provides access to the Inventor Application object.
        ''' The FirstTime flag indicates if the AddIn is loaded for the first time.
        ''' However, with the introduction of the ribbon this argument is always true.
        ''' </summary>
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate

            Dim addinName = Reflection.Assembly.GetExecutingAssembly.GetName.Name.ToString

            ' Initialize add-in members.
            _inventorApplication = addInSiteObject.Application

            ' Connect to events.
            _userInterfaceEvents = _inventorApplication.UserInterfaceManager.UserInterfaceEvents
            _userInputEvents = _inventorApplication.CommandManager.UserInputEvents
            _transactionEvents = _inventorApplication.TransactionManager.TransactionEvents
            _applicationEvents = _inventorApplication.ApplicationEvents

            ' Add event handlers.
            AddHandler _applicationEvents.OnNewDocument, AddressOf Me.Events_OnNewDocument
            AddHandler _applicationEvents.OnOpenDocument, AddressOf Me.Events_OnOpenDocument
            AddHandler _applicationEvents.OnSaveDocument, AddressOf Me.Events_OnSaveDocument
            'AddHandler _applicationEvents.OnDocumentChange, AddressOf Me.Events_PartListCreateEvent
            'AddHandler _userInputEvents.OnLinearMarkingMenu, AddressOf Me.Events_OnLinearMarkingMenu
            'AddHandler _userInputEvents.OnRadialMarkingMenu, AddressOf Me.Events_OnRadialMarkingMenu
            'AddHandler _userInputEvents.OnDrag, AddressOf Me.Events_OnDrag
            'AddHandler _transactionEvents.OnCommit, AddressOf Me.Events_OnCommit
            'AddHandler _transactionEvents.OnUndo, AddressOf Me.Events_OnUndo
            'AddHandler _transactionEvents.OnRedo, AddressOf Me.Events_OnRedo
            'AddHandler _transactionEvents.OnDelete, AddressOf Me.Events_OnDelete

            ' Register external rules path with iLogic.
            RegisterExternalRulesPathOnActivation()

#Region "Activate user interface"

            ' Add to the user interface, if it's the first time.
            ' If this add-in doesn't have a UI but runs in the background listening to events, you can delete this.

            Dim message As String = Nothing
            If firstTime Then
                Try
                    AddToUserInterface()

                    message = "Adding " & addinName & " To User Interface."
                    _inventorApplication.StatusBarText = message
                    System.Diagnostics.Debug.WriteLine("*******  " & message)
                    '  MsgBox(Message)
                Catch ex As Exception
                    MsgBox("Error" & message & vbCrLf & vbCrLf & ex.Message)
                    System.Diagnostics.Debug.WriteLine("*******  " & "Error" & message)
                End Try

            Else
                'MsgBox(addinName & " not the first time activated.")
                _inventorApplication.StatusBarText = addinName & " not the first time activated."
            End If
#End Region

        End Sub

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            _inventorApplication = Nothing
            _userInterfaceEvents = Nothing
            _userInputEvents = Nothing
            _transactionEvents = Nothing

            RemoveHandler _applicationEvents.OnOpenDocument, AddressOf Me.Events_OnOpenDocument
            RemoveHandler _applicationEvents.OnNewDocument, AddressOf Me.Events_OnNewDocument
            RemoveHandler _applicationEvents.OnSaveDocument, AddressOf Me.Events_OnSaveDocument

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()

        End Sub

        ' This property is provided to allow the AddIn to expose an API of its own to other programs.
        ' Typically, this would be done by implementing the AddIn's API interface in a class and returning that class object through this property.
        ' Typically it's not used, like in this case, and returns Nothing.
        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ' Note:tThis method is now obsolete, you should use the ControlDefinition functionality for implementing commands.
        Public Sub ExecuteCommand(ByVal commandId As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
            ' Not used.
        End Sub

        Private Sub m_uiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles _userInterfaceEvents.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub

#End Region

    End Class

End Namespace
