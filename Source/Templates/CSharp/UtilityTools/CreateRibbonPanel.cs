using Inventor;

namespace InventorAddIn.UtilityTools;

/// <summary>
/// 	Creates the ribbon panels.
/// </summary>
internal static class CreateRibbonPanel
{
	/// <summary>
	/// 	Creates a ribbon panel, or gets the panel if it already exists.
	/// </summary>
	/// <param name="panelName">
	/// 	The display name of the panel.
	/// </param>
	/// <param name="ribbonTab">
	/// 	The tab that gets the panel.
	/// </param>
	/// <returns>
	/// 	The panel.
	/// </returns>
	public static RibbonPanel GetRibbonPanel(string panelName, RibbonTab ribbonTab)
	{
		RibbonPanel? ribbonPanel = CheckForExisting.GetRibbonPanel(ribbonTab, panelName);

		if (ribbonPanel is null)
		{
			ribbonPanel = ribbonTab.RibbonPanels.Add(panelName, CheckForExisting.GetInternalName(panelName), Globals.AddInClientId);
			System.Diagnostics.Debug.WriteLine($"*******  {panelName} panel created");
		}

		return ribbonPanel;
	}
}
