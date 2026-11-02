using Autodesk.iLogic.Interfaces;
using Inventor;
using System.Reflection;

namespace InventorAddIn.UtilityTools;

/// <summary>
/// 	Runs external iLogic rules from the add-in.
/// 	Also registers the ExternalRules folder of the add-in with iLogic.
/// </summary>
internal static class RunExternaliLogicRule
{
	private const string ILogicAddInGuid = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";

	private static bool _externalRulesPathRegistered;

	/// <summary>
	/// 	Runs an external iLogic rule.
	/// </summary>
	/// <param name="externalRuleName">
	/// 	The name of the rule.
	/// </param>
	public static void RunExternalRule(string externalRuleName)
	{
		try
		{
			_ = GetILogicAutomation().RunExternalRule(Globals.InventorApplication.ActiveDocument, externalRuleName);
		}
		catch (Exception)
		{
			_ = MessageBox.Show(
				$"Error launching external rule:\n     {externalRuleName}\n\n" +
				"Ensure the iLogic rule exists, and that the\n" +
				"configuration includes the path to the rules.\n\n" +
				"See Tools tab > Options flyout button > iLogic Configuration button\n" +
				"    If needed, add an external rules folder path,\n" +
				"    and ensure that the rule is found in this folder.");
		}
	}

	/// <summary>
	/// 	Registers the ExternalRules folder of the add-in with iLogic.
	/// 	This only runs once in each session. The user can still add the folder manually if it fails.
	/// </summary>
	public static void RegisterExternalRulesPathOnActivation()
	{
		if (_externalRulesPathRegistered)
			return;

		_externalRulesPathRegistered = true;

		try
		{
			// The ExternalRules folder is next to the add-in DLL.
			string? addInDirectory = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			if (addInDirectory is null)
				return;

			string externalRulesPath = System.IO.Path.Combine(addInDirectory, "ExternalRules");
			if (System.IO.Directory.Exists(externalRulesPath) is false)
				return;

			iLogicFileOptions fileOptions = GetILogicAutomation().FileOptions;
			string[] existingPaths = fileOptions.ExternalRuleDirectories;

			if (existingPaths.Contains(externalRulesPath, StringComparer.OrdinalIgnoreCase))
				return;

			fileOptions.ExternalRuleDirectories = [.. existingPaths, externalRulesPath];
		}
		catch (Exception ex)
		{
			System.Diagnostics.Debug.WriteLine($"Could not register the external rules path: {ex.Message}");
		}
	}

	/// <summary>
	/// 	Starts the iLogic add-in and gets its automation object.
	/// </summary>
	/// <returns>
	/// 	The iLogic automation object.
	/// </returns>
	private static IiLogicAutomation GetILogicAutomation()
	{
		ApplicationAddIn iLogicAddIn = Globals.InventorApplication.ApplicationAddIns.ItemById[ILogicAddInGuid];
		iLogicAddIn.Activate();

		return (IiLogicAutomation)iLogicAddIn.Automation;
	}
}
