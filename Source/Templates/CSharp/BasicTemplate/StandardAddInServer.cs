using Inventor;
using InventorAddIn.CommandTools;
using InventorAddIn.UtilityTools;
using System.Reflection;
using System.Runtime.InteropServices;

namespace InventorAddIn;

/// <summary>
/// 	The class that Inventor loads.
/// 	Inventor and the add-in communicate through the <see cref="ApplicationAddInServer" /> members.
/// </summary>
[ProgId("InventorAddIn.StandardAddInServer"), Guid(Globals.SimpleAddInClientId)]
public class StandardAddInServer : ApplicationAddInServer
{
	// Define custom ribbon tab prefix and suffix.
	// These combine with each environment name to make ribbon tabs like "ACME Drawing Tools" and "ACME Part Tools".
	private const string RibbonTabNamePrefix = "ACME";
	private const string RibbonTabNameSuffix = "Tools";

	// Define the panels to create on each custom ribbon tab.
	private const string GeneralToolsPanelName = "General Tools";
	private const string ExternalRuleButtonsPanelName = "External Rule Buttons";

	// Inventor stops sending events when the garbage collector releases an events object.
	// So the add-in keeps each events object, and each delegate that it must unsubscribe, in a field.
	private readonly ApplicationEventsSink_OnNewDocumentEventHandler _onNewDocument;
	private readonly ApplicationEventsSink_OnOpenDocumentEventHandler _onOpenDocument;
	private readonly ApplicationEventsSink_OnSaveDocumentEventHandler _onSaveDocument;
	private readonly UserInterfaceEventsSink_OnResetRibbonInterfaceEventHandler _onResetRibbonInterface;

	private readonly Dictionary<string, (ButtonDefinition Button, ButtonDefinitionSink_OnExecuteEventHandler OnExecute)> _connectedButtons = [];
	private readonly List<RibbonTab> _customTabs = [];

	private ApplicationEvents? _applicationEvents;
	private UserInterfaceEvents? _userInterfaceEvents;
	private UserInputEvents? _userInputEvents;
	private TransactionEvents? _transactionEvents;

	/// <summary>
	/// 	Creates the event delegates.
	/// </summary>
	public StandardAddInServer()
	{
		_onNewDocument = Events_OnNewDocument;
		_onOpenDocument = Events_OnOpenDocument;
		_onSaveDocument = Events_OnSaveDocument;
		_onResetRibbonInterface = UserInterfaceEvents_OnResetRibbonInterface;
	}

	/// <summary>
	/// 	Gets the API of this add-in for other programs.
	/// 	This add-in has no API, so the value is null.
	/// </summary>
	public object? Automation => null;

	/// <summary>
	/// 	Inventor calls this method when it loads the add-in.
	/// </summary>
	/// <param name="addInSiteObject">
	/// 	Gives access to the Inventor application object.
	/// </param>
	/// <param name="firstTime">
	/// 	Tells if the add-in loads for the first time.
	/// 	With the ribbon interface, the value is always true.
	/// </param>
	public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
	{
		string? addInName = Assembly.GetExecutingAssembly().GetName().Name;

		// Initialize add-in members.
		Globals.InventorApplication = addInSiteObject.Application;

		// Connect to events.
		_userInterfaceEvents = Globals.InventorApplication.UserInterfaceManager.UserInterfaceEvents;
		_userInputEvents = Globals.InventorApplication.CommandManager.UserInputEvents;
		_transactionEvents = Globals.InventorApplication.TransactionManager.TransactionEvents;
		_applicationEvents = Globals.InventorApplication.ApplicationEvents;

		// Add event handlers.
		_userInterfaceEvents.OnResetRibbonInterface += _onResetRibbonInterface;
		_applicationEvents.OnNewDocument += _onNewDocument;
		_applicationEvents.OnOpenDocument += _onOpenDocument;
		_applicationEvents.OnSaveDocument += _onSaveDocument;

		// Add a handler to these events if the add-in needs them.
		// _userInputEvents.OnLinearMarkingMenu
		// _userInputEvents.OnRadialMarkingMenu
		// _userInputEvents.OnDrag
		// _transactionEvents.OnCommit
		// _transactionEvents.OnUndo
		// _transactionEvents.OnRedo
		// _transactionEvents.OnDelete

		// Register external rules path with iLogic.
		RunExternaliLogicRule.RegisterExternalRulesPathOnActivation();

		// Add to the user interface, if it's the first time.
		// If this add-in has no UI but runs in the background and listens to events, you can delete this.
		if (firstTime)
		{
			try
			{
				AddToUserInterface();

				string message = $"Adding {addInName} To User Interface.";
				Globals.InventorApplication.StatusBarText = message;
				System.Diagnostics.Debug.WriteLine($"*******  {message}");
			}
			catch (Exception ex)
			{
				_ = MessageBox.Show($"Could not add {addInName} to the user interface.\n\n{ex.Message}");
				System.Diagnostics.Debug.WriteLine($"*******  Could not add {addInName} to the user interface: {ex.Message}");
			}
		}
		else
		{
			Globals.InventorApplication.StatusBarText = $"{addInName} not the first time activated.";
		}
	}

	/// <summary>
	/// 	Inventor calls this method when it unloads the add-in.
	/// 	This occurs when the user unloads the add-in or when the Inventor session ends.
	/// </summary>
	public void Deactivate()
	{
		RemoveFromUserInterface();

		if (_applicationEvents is not null)
		{
			_applicationEvents.OnNewDocument -= _onNewDocument;
			_applicationEvents.OnOpenDocument -= _onOpenDocument;
			_applicationEvents.OnSaveDocument -= _onSaveDocument;
		}

		if (_userInterfaceEvents is not null)
			_userInterfaceEvents.OnResetRibbonInterface -= _onResetRibbonInterface;

		_applicationEvents = null;
		_userInterfaceEvents = null;
		_userInputEvents = null;
		_transactionEvents = null;
		Globals.InventorApplication = null!;

		GC.Collect();
		GC.WaitForPendingFinalizers();
	}

	/// <summary>
	/// 	This method is obsolete. Use a control definition to implement a command.
	/// </summary>
	/// <param name="commandId">
	/// 	Not used.
	/// </param>
	public void ExecuteCommand(int commandId)
	{
	}

	private void AddToUserInterface()
	{
		// Create the custom tab and panels for each environment.
		(RibbonPanel drawingGeneralToolsPanel, RibbonPanel drawingExternalRuleButtonsPanel) = CreateCustomTab("Drawing");
		(RibbonPanel assemblyGeneralToolsPanel, _) = CreateCustomTab("Assembly");
		(RibbonPanel partGeneralToolsPanel, _) = CreateCustomTab("Part");

		// Add this button to the General Tools panel of 3 different environments.
		ConnectButton(ApiHelp.CreateButton(drawingGeneralToolsPanel, true, false), _ => ApiHelp.RunCommandCode());
		ConnectButton(ApiHelp.CreateButton(assemblyGeneralToolsPanel, true, false), _ => ApiHelp.RunCommandCode());
		ConnectButton(ApiHelp.CreateButton(partGeneralToolsPanel, true, false), _ => ApiHelp.RunCommandCode());

		// Add this button to just one panel.
		ConnectButton(DetectInventorTheme.CreateButton(drawingExternalRuleButtonsPanel, true, false), _ => DetectInventorTheme.RunExternalRule());
	}

	/// <summary>
	/// 	Creates a custom tab and its panels in the ribbon of an environment.
	/// </summary>
	/// <remarks>
	/// 	Ex. the environment "Drawing", the prefix "ACME", and the suffix "Tools" create a tab with the name "ACME Drawing Tools".
	/// </remarks>
	/// <param name="environment">
	/// 	The ribbon name of the environment, ex. "Drawing".
	/// </param>
	/// <returns>
	/// 	The panels on the tab.
	/// </returns>
	private (RibbonPanel GeneralTools, RibbonPanel ExternalRuleButtons) CreateCustomTab(string environment)
	{
		Ribbon ribbon = Globals.InventorApplication.UserInterfaceManager.Ribbons[environment];

		// The custom tab goes after the last visible tab.
		string lastTabName = string.Empty;
		foreach (RibbonTab tab in ribbon.RibbonTabs)
		{
			if (tab.Visible)
				lastTabName = tab.InternalName;
		}

		string newTabName = $"{RibbonTabNamePrefix} {ribbon.InternalName} {RibbonTabNameSuffix}";
		RibbonTab customTab = CreateRibbonTab.GetRibbonTab(newTabName, ribbon, lastTabName);
		_customTabs.Add(customTab);

		return (
			CreateRibbonPanel.GetRibbonPanel(GeneralToolsPanelName, customTab),
			CreateRibbonPanel.GetRibbonPanel(ExternalRuleButtonsPanelName, customTab));
	}

	/// <summary>
	/// 	Connects the click event of a button.
	/// </summary>
	/// <remarks>
	/// 	One button can be on many tabs, and a ribbon reset creates the buttons again.
	/// 	Each call gives the same button definition, so only the first call connects the event.
	/// 	A second connection would run the command twice for one click.
	/// </remarks>
	/// <param name="button">
	/// 	The button definition.
	/// </param>
	/// <param name="onExecute">
	/// 	The code to run when the user clicks the button.
	/// </param>
	private void ConnectButton(ButtonDefinition button, ButtonDefinitionSink_OnExecuteEventHandler onExecute)
	{
		if (_connectedButtons.ContainsKey(button.InternalName))
			return;

		button.OnExecute += onExecute;
		_connectedButtons.Add(button.InternalName, (button, onExecute));
	}

	private void RemoveFromUserInterface()
	{
		foreach ((ButtonDefinition button, ButtonDefinitionSink_OnExecuteEventHandler onExecute) in _connectedButtons.Values)
			button.OnExecute -= onExecute;

		_connectedButtons.Clear();

		foreach (RibbonTab customTab in _customTabs)
		{
			try
			{
				customTab.Delete();
			}
			catch (Exception ex)
			{
				// A ribbon reset already deleted the tab.
				System.Diagnostics.Debug.WriteLine($"Could not delete a ribbon tab: {ex.Message}");
			}
		}

		_customTabs.Clear();
	}

	private void UserInterfaceEvents_OnResetRibbonInterface(NameValueMap context)
	{
		// The ribbon was reset, so add back the user interface of the add-in.
		RemoveFromUserInterface();
		AddToUserInterface();
	}

	private void Events_OnNewDocument(_Document documentObject, EventTimingEnum beforeOrAfter, NameValueMap context, out HandlingCodeEnum handlingCode)
	{
		handlingCode = HandlingCodeEnum.kEventNotHandled;

		if (documentObject is null)
			return;

		if (documentObject.DocumentType == DocumentTypeEnum.kDrawingDocumentObject && beforeOrAfter == EventTimingEnum.kAfter)
		{
			// Run a command that has no button.
			// Uncomment the following line to see the Welcome command run on this event.
			// Welcome.AddinHello(documentObject.FullFileName, beforeOrAfter.ToString());
		}
	}

	private void Events_OnOpenDocument(_Document documentObject, string fullDocumentName, EventTimingEnum beforeOrAfter, NameValueMap context, out HandlingCodeEnum handlingCode)
	{
		handlingCode = HandlingCodeEnum.kEventNotHandled;

		if (documentObject is null)
			return;

		// Only trigger on the active file.
		if (ReferenceEquals(documentObject, Globals.InventorApplication.ActiveDocument) is false)
			return;

		if (documentObject.DocumentType != DocumentTypeEnum.kDrawingDocumentObject)
			return;

		// Run this after a document is open.
		if (beforeOrAfter == EventTimingEnum.kAfter)
		{
			// Run a command that has no button.
			// Uncomment the following line to see the Welcome command run on this event.
			// Welcome.AddinHello(documentObject.FullFileName, beforeOrAfter.ToString());
		}
	}

	private void Events_OnSaveDocument(_Document documentObject, EventTimingEnum beforeOrAfter, NameValueMap context, out HandlingCodeEnum handlingCode)
	{
		handlingCode = HandlingCodeEnum.kEventNotHandled;

		if (documentObject.DocumentType == DocumentTypeEnum.kDrawingDocumentObject && beforeOrAfter == EventTimingEnum.kBefore)
		{
			// Run a command that has no button.
			// Uncomment the following line to see the Welcome command run on this event.
			// Welcome.AddinHello(documentObject.FullFileName, beforeOrAfter.ToString());
		}
	}
}
