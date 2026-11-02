using Inventor;

namespace InventorAddIn.UtilityTools;

/// <summary>
/// 	Creates the ribbon tabs.
/// </summary>
internal static class CreateRibbonTab
{
	/// <summary>
	/// 	Creates a ribbon tab, or gets the tab if it already exists.
	/// </summary>
	/// <param name="tabName">
	/// 	The display name of the tab.
	/// </param>
	/// <param name="ribbon">
	/// 	The ribbon that gets the tab.
	/// </param>
	/// <param name="tabNextTo">
	/// 	The internal name of the tab that the new tab goes after.
	/// </param>
	/// <returns>
	/// 	The tab.
	/// </returns>
	public static RibbonTab GetRibbonTab(string tabName, Ribbon ribbon, string tabNextTo)
	{
		RibbonTab? ribbonTab = CheckForExisting.GetRibbonTab(ribbon, tabName);

		if (ribbonTab is null)
		{
			ribbonTab = ribbon.RibbonTabs.Add(tabName, CheckForExisting.GetInternalName(tabName), Globals.AddInClientId, tabNextTo, false);
			System.Diagnostics.Debug.WriteLine($"*******  {tabName} tab created");
		}

		return ribbonTab;
	}
}
