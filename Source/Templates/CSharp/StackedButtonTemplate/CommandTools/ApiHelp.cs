using Inventor;
using InventorAddIn.Properties;
using InventorAddIn.UtilityTools;
using System.Diagnostics;

namespace InventorAddIn.CommandTools;

/// <summary>
/// 	Opens the API help.
/// 	<see cref="CreateButton" /> defines the label, the icons, and the tool tip of the button.
/// </summary>
internal static class ApiHelp
{
	/// <summary>
	/// 	Creates the button.
	/// </summary>
	/// <param name="ribbonPanel">
	/// 	The panel that gets the button.
	/// </param>
	/// <param name="useLargeIcon">
	/// 	True to show the large icon, and false to show the standard icon.
	/// </param>
	/// <param name="isInButtonStack">
	/// 	True if the button goes in a button stack.
	/// </param>
	/// <returns>
	/// 	The button definition.
	/// </returns>
	public static ButtonDefinition CreateButton(RibbonPanel ribbonPanel, bool useLargeIcon, bool isInButtonStack)
	{
		// Get the images to use for the button.
		stdole.IPictureDisp largeIcon = PictureDispConverter.ToIPictureDisp(Resources.Cat_32);
		stdole.IPictureDisp standardIcon = PictureDispConverter.ToIPictureDisp(Resources.Cat_16);
		stdole.IPictureDisp? toolTipImage = null;

		// This is the text the user sees on the button.
		string buttonLabel = "API\nHelp";

		// Text that displays when the user hovers over the button.
		string toolTipSimple = "Opens the API Help";
		string? toolTipExpanded = null;

		// Set to true to use a progressive tool tip, and false to use a simple tool tip.
		bool useProgressToolTip = true;

		// Only used if 'useProgressToolTip' is true.
		if (useProgressToolTip)
		{
			toolTipImage = PictureDispConverter.ToIPictureDisp(Resources.API_ToolTip);

			toolTipExpanded =
				"\u2022 This tool opens the API help *.chm file in a separate window\n" +
				"\u2022 Line2\n" +
				"\u2022 Line3\n" +
				"\u2022 Line4";
		}

		return CreateButtonDefinition.Create(
			ribbonPanel,
			useLargeIcon,
			isInButtonStack,
			useProgressToolTip,
			buttonLabel,
			toolTipSimple,
			toolTipExpanded,
			standardIcon,
			largeIcon,
			toolTipImage);
	}

	/// <summary>
	/// 	Opens the Inventor API help file.
	/// </summary>
	public static void RunCommandCode()
	{
		SoftwareVersion softwareVersion = Globals.InventorApplication.SoftwareVersion;
		string version = softwareVersion.DisplayVersion.Split('.')[0];
		string fileName = $@"C:\Users\Public\Documents\Autodesk\Inventor {version}\Local Help\ADMAPI_{softwareVersion.Major}_0.chm";

		if (System.IO.File.Exists(fileName))
			_ = Process.Start(new ProcessStartInfo(fileName) { UseShellExecute = true });
		else
			_ = MessageBox.Show($"File Does Not Exist\n{fileName}");
	}
}
