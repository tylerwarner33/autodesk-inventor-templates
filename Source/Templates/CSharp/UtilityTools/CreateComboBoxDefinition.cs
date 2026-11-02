using Inventor;

namespace InventorAddIn.UtilityTools;

/// <summary>
/// 	Creates the combo box definitions.
/// </summary>
internal static class CreateComboBoxDefinition
{
	/// <summary>
	/// 	Creates a combo box, or gets the combo box if it already exists.
	/// </summary>
	/// <param name="ribbonPanel">
	/// 	The panel that gets the combo box.
	/// </param>
	/// <param name="isInButtonStack">
	/// 	True if the combo box goes in a button stack. The panel then does not get the combo box directly.
	/// </param>
	/// <param name="isInSlideOut">
	/// 	True to put the combo box in the slide-out area of the panel.
	/// </param>
	/// <param name="label">
	/// 	The text of the combo box.
	/// </param>
	/// <param name="dropDownWidth">
	/// 	The width of the drop-down list.
	/// </param>
	/// <param name="toolTipSimple">
	/// 	The simple tool tip.
	/// </param>
	/// <returns>
	/// 	The combo box definition.
	/// </returns>
	public static ComboBoxDefinition Create(
		RibbonPanel ribbonPanel,
		bool isInButtonStack,
		bool isInSlideOut,
		string label,
		int dropDownWidth,
		string toolTipSimple)
	{
		string labelNoSpaces = label.Replace("\r\n", "_").Replace("\n", "_").Replace("\r", "_").Replace(" ", "_");
		string internalName = $"{labelNoSpaces}_{Globals.AddInClientId}_ComboBox_InternalName";

		ComboBoxDefinition? comboBoxDefinition = CheckForExisting.GetComboBoxDefinition(internalName);

		if (comboBoxDefinition is not null)
			return comboBoxDefinition;

		comboBoxDefinition = Globals.InventorApplication.CommandManager.ControlDefinitions.AddComboBoxDefinition(
			label,
			internalName,
			CommandTypesEnum.kEditMaskCmdType,
			dropDownWidth,
			Globals.AddInClientId,
			$"{labelNoSpaces} ComboBox",
			toolTipSimple);

		if (isInButtonStack is false)
		{
			if (isInSlideOut)
				_ = ribbonPanel.SlideoutControls.AddComboBox(comboBoxDefinition);
			else
				_ = ribbonPanel.CommandControls.AddComboBox(comboBoxDefinition);
		}

		return comboBoxDefinition;
	}
}
