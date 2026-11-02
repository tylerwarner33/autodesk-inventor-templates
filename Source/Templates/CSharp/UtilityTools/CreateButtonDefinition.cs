using Inventor;

namespace InventorAddIn.UtilityTools;

/// <summary>
/// 	Creates the button definitions.
/// </summary>
internal static class CreateButtonDefinition
{
	/// <summary>
	/// 	Creates a button, or gets the button if it already exists.
	/// </summary>
	/// <param name="ribbonPanel">
	/// 	The panel that gets the button.
	/// </param>
	/// <param name="useLargeIcon">
	/// 	True to show the large icon, and false to show the standard icon.
	/// </param>
	/// <param name="isInButtonStack">
	/// 	True if the button goes in a button stack. The panel then does not get the button directly.
	/// </param>
	/// <param name="useProgressToolTip">
	/// 	True to use a progressive tool tip, and false to use a simple tool tip.
	/// </param>
	/// <param name="buttonLabel">
	/// 	The text on the button. A line feed splits the text on a large button.
	/// </param>
	/// <param name="toolTipSimple">
	/// 	The simple tool tip.
	/// </param>
	/// <param name="toolTipExpanded">
	/// 	The expanded text of the progressive tool tip.
	/// </param>
	/// <param name="standardIcon">
	/// 	The 16 pixel icon.
	/// </param>
	/// <param name="largeIcon">
	/// 	The 32 pixel icon.
	/// </param>
	/// <param name="toolTipImage">
	/// 	The image of the progressive tool tip.
	/// </param>
	/// <returns>
	/// 	The button definition.
	/// </returns>
	public static ButtonDefinition Create(
		RibbonPanel ribbonPanel,
		bool useLargeIcon,
		bool isInButtonStack,
		bool useProgressToolTip,
		string buttonLabel,
		string toolTipSimple,
		string? toolTipExpanded,
		stdole.IPictureDisp standardIcon,
		stdole.IPictureDisp largeIcon,
		stdole.IPictureDisp? toolTipImage)
	{
		string buttonNameNoSpaces = RemoveLineBreaks(buttonLabel, "_").Replace(" ", "_");
		string internalName = $"{buttonNameNoSpaces}_{Globals.AddInClientId}_Button_InternalName";
		string description = $"{buttonNameNoSpaces} Button";

		// Only a large icon wraps the label to 2 rows.
		if (useLargeIcon is false)
			buttonLabel = RemoveLineBreaks(buttonLabel, " ");

		ButtonDefinition? buttonDefinition = CheckForExisting.GetButtonDefinition(internalName);

		if (buttonDefinition is null)
		{
			buttonDefinition = Globals.InventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition(
				buttonLabel,
				internalName,
				CommandTypesEnum.kShapeEditCmdType,
				Globals.AddInClientId,
				description,
				toolTipSimple);
			System.Diagnostics.Debug.WriteLine($"*******  {internalName} button created");
		}

		// Icons for the button.
		buttonDefinition.LargeIcon = largeIcon;
		buttonDefinition.StandardIcon = standardIcon;

		if (useProgressToolTip)
		{
			ProgressiveToolTip toolTip = buttonDefinition.ProgressiveToolTip;
			toolTip.IsProgressive = true;
			toolTip.Title = buttonLabel;
			toolTip.Description = description;
			toolTip.ExpandedDescription = toolTipExpanded;
			toolTip.Image = toolTipImage;
		}

		// A button in a button stack goes on the panel with the stack.
		if (isInButtonStack is false)
		{
			_ = ribbonPanel.CommandControls.AddButton(buttonDefinition, useLargeIcon, true);
			System.Diagnostics.Debug.WriteLine($"*******  {internalName} button added to {ribbonPanel.DisplayName}");
		}

		return buttonDefinition;
	}

	/// <summary>
	/// 	Replaces each line break in a text.
	/// </summary>
	/// <param name="text">
	/// 	The text.
	/// </param>
	/// <param name="replacement">
	/// 	The text that replaces a line break.
	/// </param>
	/// <returns>
	/// 	The text with no line breaks.
	/// </returns>
	private static string RemoveLineBreaks(string text, string replacement) => text.Replace("\r\n", replacement).Replace("\n", replacement).Replace("\r", replacement);
}
