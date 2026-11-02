using Inventor;
using InventorAddIn.Properties;
using InventorAddIn.UtilityTools;

namespace InventorAddIn.CommandTools;

/// <summary>
/// 	Creates a sketched symbol: Circle Symbol.
/// 	<see cref="CreateButton" /> defines the label, the icons, and the tool tip of the button.
/// </summary>
internal static class CreateSketchedCircle
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
		stdole.IPictureDisp largeIcon = PictureDispConverter.ToIPictureDisp(Resources.One_32);
		stdole.IPictureDisp standardIcon = PictureDispConverter.ToIPictureDisp(Resources.One_16);
		stdole.IPictureDisp? toolTipImage = null;

		// This is the text the user sees on the button.
		string buttonLabel = "New\nCircle Symbol";

		// Text that displays when the user hovers over the button.
		string toolTipSimple = "Create a new sketched symbol";
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
	/// 	Creates the sketched symbol and inserts the symbol on the active sheet.
	/// </summary>
	public static void RunCommandCode()
	{
		DrawingDocument drawingDocument = (DrawingDocument)Globals.InventorApplication.ActiveDocument;
		TransientGeometry transientGeometry = Globals.InventorApplication.TransientGeometry;
		string symbolName = "Circle Symbol";
		SketchedSymbolDefinition sketchedSymbolDefinition;

		try
		{
			// Try to get it.
			sketchedSymbolDefinition = drawingDocument.SketchedSymbolDefinitions[symbolName];
		}
		catch (Exception)
		{
			// Add it if not found.
			sketchedSymbolDefinition = drawingDocument.SketchedSymbolDefinitions.Add(symbolName);
		}

		// Open the sketch in edit mode.
		sketchedSymbolDefinition.Edit(out DrawingSketch drawingSketch);
		_ = drawingSketch.SketchCircles.AddByCenterRadius(transientGeometry.CreatePoint2d(5, 5), 2);

		// Quit the edit mode.
		sketchedSymbolDefinition.ExitEdit(true);

		// Insert the sketched symbol on the active sheet.
		_ = drawingDocument.ActiveSheet.SketchedSymbols.Add(symbolName, transientGeometry.CreatePoint2d(5, 5));
	}
}
