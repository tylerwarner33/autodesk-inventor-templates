Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Inventor

Imports Microsoft.Win32


'Define global variables
Public Module Globals

    ' Inventor application object.
    Public g_inventorApplication As Inventor.Application

    'Unique ID for this add-in.  
    Public Const g_simpleAddInClientID As String = "{{guid}}"
    Public Const g_addInClientID As String = "{" & g_simpleAddInClientID & "}"


End Module

Namespace {{ProjectName}}
    <ProgIdAttribute("{{ProjectName}}.StandardAddInServer"),
    GuidAttribute(g_simpleAddInClientID)>
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        'Application events
        Private WithEvents InventorApplicationEvents As ApplicationEvents
        Private WithEvents UiEvents As UserInterfaceEvents
        Private WithEvents UserInputEvents As UserInputEvents
        Private WithEvents InvTransactionEvents As TransactionEvents

        'Define the command buttons
        Private WithEvents API_Help_Button As ButtonDefinition
        Private WithEvents Detect_Dark_Light_Theme_Button As ButtonDefinition


        Private Sub AddToUserInterface()

#Region "Setup Information"
            'create a object collection and Control Definition array to be used with button stacks
            Dim buttonObjectCollection As ObjectCollection = g_inventorApplication.TransientObjects.CreateObjectCollection
            Dim ctrlDef(0 To 99) As ControlDefinition

            'create list of environments to cycle through
            Dim EnvironmentList As New List(Of String)({"Drawing", "Assembly", "Part"})

            'define custom roibbon tab prefix and suffix
            'this will be combined with the EnvironmentList to create ribbon tabs like "ACME Drawing Tools", "ACME Part Tools", etc.
            ' Dim RibbonTabName_Prefix As String = "ACME"
            Dim RibbonTabName_Prefix As String = "ACME"
            Dim RibbonTabName_Sufffix As String = "Tools"

            'create list of panels to create on the custom ribbon tab
            Dim CustomPanelList As New List(Of String)({"General Tools", "External Rule Buttons"})
#End Region

#Region "Create Custom Tabs and Panels for each environment in the list"
            'Create Custom Tabs and Panels for each environment in the list
            'Example:
            ''      Environment = "Drawing"
            ''      Prefix = 'ACME'
            ''      Suffix = "Tools"
            ''      will create a custom tab called: ACME Drawing Tools

            Dim AllRibbons As Ribbons = g_inventorApplication.UserInterfaceManager.Ribbons
            Dim VisibleTabs_Collection As ObjectCollection = g_inventorApplication.TransientObjects.CreateObjectCollection

            Dim CustomDrawingTab As RibbonTab = Nothing
            Dim CustomAssemblyTab As RibbonTab = Nothing
            Dim CustomPartTab As RibbonTab = Nothing

            Dim Drawing_GeneralToolsPanel As RibbonPanel = Nothing
            Dim Drawing_ExternalRuleButtonsPanel As RibbonPanel = Nothing
            Dim Assembly_GeneralToolsPanel As RibbonPanel = Nothing
            Dim Assembly_ExternalRuleButtonsPanel As RibbonPanel = Nothing
            Dim Part_GeneralToolsPanel As RibbonPanel = Nothing
            Dim Part_ExternalRuleButtonsPanel As RibbonPanel = Nothing


            'iterate through the list
            For i = 0 To EnvironmentList.Count - 1

                'get the ribbon for each environment ( example: Drawing, Assembly, Part) 
                Dim MyRibbon As Ribbon = AllRibbons.Item(EnvironmentList.Item(i))

                'make sure the collection is empty
                VisibleTabs_Collection.Clear()

                'collect only the visible tabs in the ribbon
                For Each Ribbon_Tab In MyRibbon.RibbonTabs
                    If Ribbon_Tab.Visible = False Then Continue For
                    VisibleTabs_Collection.Add(Ribbon_Tab)
                Next

                'get count of tabs in collection
                Dim iCount As Integer = VisibleTabs_Collection.Count

                'get the last tab in the collection
                Dim LastTab As String = VisibleTabs_Collection.Item(iCount).InternalName

                'create new tab name
                Dim NewTabName As String = RibbonTabName_Prefix & " " & MyRibbon.InternalName & " " & RibbonTabName_Sufffix

                'create new tab
                Dim CustomTab As RibbonTab = CreateRibbonTab.GetRibbon_Tab(NewTabName, MyRibbon, LastTab)

                'create the tab and panels for each environment
                If i = 0 Then
                    CustomDrawingTab = CustomTab
                    Drawing_GeneralToolsPanel = CreateRibbonPanel.GetRibbonPanel(CustomPanelList.Item(0), CustomTab)
                    Drawing_ExternalRuleButtonsPanel = CreateRibbonPanel.GetRibbonPanel(CustomPanelList.Item(1), CustomTab)
                ElseIf i = 1 Then
                    CustomAssemblyTab = CustomTab
                    Assembly_GeneralToolsPanel = CreateRibbonPanel.GetRibbonPanel(CustomPanelList.Item(0), CustomTab)
                    Assembly_ExternalRuleButtonsPanel = CreateRibbonPanel.GetRibbonPanel(CustomPanelList.Item(1), CustomTab)
                ElseIf i = 2 Then
                    CustomPartTab = CustomTab
                    Part_GeneralToolsPanel = CreateRibbonPanel.GetRibbonPanel(CustomPanelList.Item(0), CustomTab)
                    Part_ExternalRuleButtonsPanel = CreateRibbonPanel.GetRibbonPanel(CustomPanelList.Item(1), CustomTab)
                End If

            Next
#End Region

#Region "Create Single Buttons"

            'add this button to General Tools tab of 3 different environments 
            API_Help_Button = API_Help.CreateButton("Drawing", CustomDrawingTab, Drawing_GeneralToolsPanel, True, False)
            API_Help_Button = API_Help.CreateButton("Assembly", CustomAssemblyTab, Assembly_GeneralToolsPanel, True, False)
            API_Help_Button = API_Help.CreateButton("Part", CustomPartTab, Part_GeneralToolsPanel, True, False)

            'add button to just one tab
            Detect_Dark_Light_Theme_Button = Detect_Dark_Light_Theme.CreateButton("Drawing", CustomDrawingTab, Drawing_ExternalRuleButtonsPanel, True, False)

#End Region


        End Sub


#Region "Button 'on click' Events"
        'add button click events to this region

        Private Sub API_Help_Button_OnExecute(Context As NameValueMap) Handles API_Help_Button.OnExecute
            API_Help.RunCommandCode()
        End Sub

        Private Sub Detect_Dark_Light_Theme_Button_OnExecute(Context As NameValueMap) Handles Detect_Dark_Light_Theme_Button.OnExecute
            Detect_Dark_Light_Theme.Run_ExternalRule()
        End Sub
#End Region

#Region "Applicaton Events"

        Private Sub Events_OnNewDocument(ByVal DocumentObject As Inventor._Document,
                                             ByVal BeforeOrAfter As Inventor.EventTimingEnum,
                                             ByVal Context As Inventor.NameValueMap,
                                             ByRef HandlingCode As Inventor.HandlingCodeEnum)

            If DocumentObject Is Nothing Then Exit Sub

            If DocumentObject.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                If BeforeOrAfter = EventTimingEnum.kAfter Then
                    'run a command that has no button
                    'uncomment the following line to see the Welcome command module run on this event
                    'Welcome.Addin_Hello(DocumentObject.FullFileName, EventTimingEnum.kAfter.ToString)

                End If
            End If


        End Sub
        Private Sub Events_OnOpenDocument(DocumentObject As _Document, FullDocumentName As String,
                                    BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)

            If DocumentObject Is Nothing Then Exit Sub
            If Not DocumentObject Is g_inventorApplication.ActiveDocument Then Exit Sub 'only trigger on the active file
            If Not DocumentObject.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then Exit Sub

            'run this after a document is open
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                'run a command that has no button
                'uncomment the following line to see the Welcome command module run on this event
                'Welcome.Addin_Hello(DocumentObject.FullFileName, EventTimingEnum.kAfter.ToString)

            End If

        End Sub

        Private Sub Events_OnSaveDocument(DocumentObject As _Document,
                                    BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)


            If DocumentObject.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                If BeforeOrAfter = EventTimingEnum.kBefore Then
                    'run a command that has no button
                    'uncomment the following line to see the Welcome command module run on this event
                    'Welcome.Addin_Hello(DocumentObject.FullFileName, EventTimingEnum.kAfter.ToString)
                End If
            End If

        End Sub

#End Region

#Region "ApplicationAddInServer Members"

        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate

            Dim AddinName = Reflection.Assembly.GetExecutingAssembly.GetName.Name.ToString

            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' Connect to events 
            UiEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents
            UserInputEvents = g_inventorApplication.CommandManager.UserInputEvents
            InvTransactionEvents = g_inventorApplication.TransactionManager.TransactionEvents
            InventorApplicationEvents = g_inventorApplication.ApplicationEvents

            'add event handlers
            AddHandler InventorApplicationEvents.OnNewDocument, AddressOf Me.Events_OnNewDocument
            AddHandler InventorApplicationEvents.OnOpenDocument, AddressOf Me.Events_OnOpenDocument
            AddHandler InventorApplicationEvents.OnSaveDocument, AddressOf Me.Events_OnSaveDocument
            'AddHandler UserInputEvents.OnLinearMarkingMenu, AddressOf Me.Events_OnLinearMarkingMenu
            'AddHandler UserInputEvents.OnRadialMarkingMenu, AddressOf Me.Events_OnRadialMarkingMenu
            'AddHandler InventorApplicationEvents.OnDocumentChange, AddressOf Me.Events_PartListCreateEvent
            'AddHandler UserInputEvents.OnDrag, AddressOf Me.Events_OnDrag
            'AddHandler InvTransactionEvents.OnCommit, AddressOf Me.Events_OnCommit
            'AddHandler InvTransactionEvents.OnUndo, AddressOf Me.Events_OnUndo
            'AddHandler InvTransactionEvents.OnRedo, AddressOf Me.Events_OnRedo
            'AddHandler InvTransactionEvents.OnDelete, AddressOf Me.Events_OnDelete


#Region "Activate user interface"
            ' Add to the user interface, if it's the first time.
            ' If this add-in doesn't have a UI but runs in the background listening
            ' to events, you can delete this.

            Dim Message As String = Nothing
            If firstTime Then
                Try
                    AddToUserInterface()

                    Message = "Adding " & AddinName & " To User Interface."
                    g_inventorApplication.StatusBarText = Message
                    System.Diagnostics.Debug.WriteLine("*******  " & Message)
                    '  MsgBox(Message)
                Catch ex As Exception
                    MsgBox("Error" & Message & vbCrLf & vbCrLf & ex.Message)
                    System.Diagnostics.Debug.WriteLine("*******  " & "Error" & Message)
                End Try

            Else
                'MsgBox(AddinName & " not the first time activated.")
                g_inventorApplication.StatusBarText = AddinName & " not the first time activated."
            End If
#End Region

        End Sub

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            g_inventorApplication = Nothing
            UiEvents = Nothing
            UserInputEvents = Nothing
            InvTransactionEvents = Nothing

            RemoveHandler InventorApplicationEvents.OnOpenDocument, AddressOf Me.Events_OnOpenDocument
            RemoveHandler InventorApplicationEvents.OnNewDocument, AddressOf Me.Events_OnNewDocument
            RemoveHandler InventorApplicationEvents.OnSaveDocument, AddressOf Me.Events_OnSaveDocument

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        ' This property is provided to allow the AddIn to expose an API of its own to other 
        ' programs. Typically, this  would be done by implementing the AddIn's API
        ' interface in a class and returning that class object through this property.
        ' Typically it's not used, like in this case, and returns Nothing.
        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ' Note:this method is now obsolete, you should use the 
        ' ControlDefinition functionality for implementing commands.
        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
            ' Not used.
        End Sub

        Private Sub m_uiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles UiEvents.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub

#End Region

    End Class


End Namespace


