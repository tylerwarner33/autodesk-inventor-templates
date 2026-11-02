using Inventor;

namespace InventorAddIn.UtilityTools;

/// <summary>
/// 	Finds the ribbon items and the control definitions that already exist.
/// </summary>
/// <remarks>
/// 	An Inventor collection throws when it has no item with the name.
/// 	So each method catches the exception and returns null.
/// </remarks>
internal static class CheckForExisting
{
	/// <summary>
	/// 	Finds a button definition.
	/// </summary>
	/// <param name="internalName">
	/// 	The internal name of the button definition.
	/// </param>
	/// <returns>
	/// 	The button definition, or null if it does not exist.
	/// </returns>
	public static ButtonDefinition? GetButtonDefinition(string internalName)
	{
		try
		{
			ButtonDefinition? buttonDefinition = Globals.InventorApplication.CommandManager.ControlDefinitions[internalName] as ButtonDefinition;
			System.Diagnostics.Debug.WriteLine($"*******  {internalName} button exists");
			return buttonDefinition;
		}
		catch (Exception)
		{
			return null;
		}
	}

	/// <summary>
	/// 	Finds a combo box definition.
	/// </summary>
	/// <param name="internalName">
	/// 	The internal name of the combo box definition.
	/// </param>
	/// <returns>
	/// 	The combo box definition, or null if it does not exist.
	/// </returns>
	public static ComboBoxDefinition? GetComboBoxDefinition(string internalName)
	{
		try
		{
			ComboBoxDefinition? comboBoxDefinition = Globals.InventorApplication.CommandManager.ControlDefinitions[internalName] as ComboBoxDefinition;
			System.Diagnostics.Debug.WriteLine($"*******  {internalName} combo box exists");
			return comboBoxDefinition;
		}
		catch (Exception)
		{
			return null;
		}
	}

	/// <summary>
	/// 	Finds a ribbon panel.
	/// </summary>
	/// <param name="ribbonTab">
	/// 	The tab to look in.
	/// </param>
	/// <param name="panelName">
	/// 	The display name of the panel.
	/// </param>
	/// <returns>
	/// 	The panel, or null if it does not exist.
	/// </returns>
	public static RibbonPanel? GetRibbonPanel(RibbonTab ribbonTab, string panelName)
	{
		try
		{
			RibbonPanel ribbonPanel = ribbonTab.RibbonPanels[GetInternalName(panelName)];
			System.Diagnostics.Debug.WriteLine($"*******  {panelName} panel exists");
			return ribbonPanel;
		}
		catch (Exception)
		{
			return null;
		}
	}

	/// <summary>
	/// 	Finds a ribbon tab.
	/// </summary>
	/// <param name="ribbon">
	/// 	The ribbon to look in.
	/// </param>
	/// <param name="tabName">
	/// 	The display name of the tab.
	/// </param>
	/// <returns>
	/// 	The tab, or null if it does not exist.
	/// </returns>
	public static RibbonTab? GetRibbonTab(Ribbon ribbon, string tabName)
	{
		try
		{
			RibbonTab ribbonTab = ribbon.RibbonTabs[GetInternalName(tabName)];
			System.Diagnostics.Debug.WriteLine($"*******  {tabName} tab exists");
			return ribbonTab;
		}
		catch (Exception)
		{
			return null;
		}
	}

	/// <summary>
	/// 	Makes the internal name of a ribbon tab or a ribbon panel.
	/// </summary>
	/// <param name="displayName">
	/// 	The display name of the tab or the panel.
	/// </param>
	/// <returns>
	/// 	The internal name.
	/// </returns>
	public static string GetInternalName(string displayName) => $"id_{displayName.Replace(" ", "_")}";
}
