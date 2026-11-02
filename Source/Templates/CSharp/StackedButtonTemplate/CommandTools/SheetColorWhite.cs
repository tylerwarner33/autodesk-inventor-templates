using Inventor;
using InventorAddIn.Properties;
using InventorAddIn.UtilityTools;

namespace InventorAddIn.CommandTools;

/// <summary>
/// 	Sets the drawing sheet color: White.
/// 	<see cref="CreateButton" /> defines the label, the icons, and the tool tip of the button.
/// </summary>
internal static class SheetColorWhite
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
		stdole.IPictureDisp largeIcon = PictureDispConverter.ToIPictureDisp(Resources.WhiteColor_32);
		stdole.IPictureDisp standardIcon = PictureDispConverter.ToIPictureDisp(Resources.WhiteColor_16);
		stdole.IPictureDisp? toolTipImage = null;

		// This is the text the user sees on the button.
		string buttonLabel = "Sheet Color\nWhite";

		// Text that displays when the user hovers over the button.
		string toolTipSimple = "Set the color of the drawing sheet";
		string? toolTipExpanded = null;

		// Set to true to use a progressive tool tip, and false to use a simple tool tip.
		bool useProgressToolTip = false;

		// Only used if 'useProgressToolTip' is true.
		if (useProgressToolTip)
		{
			toolTipImage = PictureDispConverter.ToIPictureDisp(Resources.AU_ToolTip);

			toolTipExpanded =
				"\u2022 Line1\n" +
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
	/// 	Sets the color of the drawing sheet.
	/// </summary>
	public static void RunCommandCode()
	{
		DrawingDocument drawingDocument = (DrawingDocument)Globals.InventorApplication.ActiveDocument;
		TransientObjects transientObjects = Globals.InventorApplication.TransientObjects;
		drawingDocument.SheetSettings.SheetColor = transientObjects.CreateColor(255, 255, 255);
	}
}
