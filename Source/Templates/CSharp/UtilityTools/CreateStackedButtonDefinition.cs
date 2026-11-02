using Inventor;

namespace InventorAddIn.UtilityTools;

/// <summary>
/// 	Creates the button stacks.
/// </summary>
internal static class CreateStackedButtonDefinition
{
	/// <summary>
	/// 	Creates a simple stack of buttons.
	/// 	See the 'Switch' control in the Windows panel of the View tab.
	/// </summary>
	/// <param name="ribbonPanel">
	/// 	The panel that gets the stack.
	/// </param>
	/// <param name="collection">
	/// 	The buttons in the stack.
	/// </param>
	/// <param name="defaultButton">
	/// 	The button that shows on the panel.
	/// </param>
	/// <param name="useLargeIcon">
	/// 	True to show the large icon, and false to show the standard icon.
	/// </param>
	public static void CreatePopUpButtonDefinition(RibbonPanel ribbonPanel, ObjectCollection collection, ButtonDefinition defaultButton, bool useLargeIcon)
	{
		try
		{
			_ = ribbonPanel.CommandControls.AddPopup(defaultButton, collection, useLargeIcon, true, "", false);
		}
		catch (Exception ex)
		{
			System.Diagnostics.Debug.WriteLine($"Error in {nameof(CreatePopUpButtonDefinition)}: {ribbonPanel.DisplayName}, {defaultButton.DisplayName}");
			System.Diagnostics.Debug.WriteLine(ex.Message);
		}
	}

	/// <summary>
	/// 	Creates a stack of buttons with check boxes that the user can toggle on and off.
	/// 	See the 'Object Visibility' control in the Visibility panel of the View tab.
	/// </summary>
	/// <param name="ribbonPanel">
	/// 	The panel that gets the stack.
	/// </param>
	/// <param name="collection">
	/// 	The buttons in the stack.
	/// </param>
	/// <param name="defaultButton">
	/// 	The button that shows on the panel.
	/// </param>
	/// <param name="useLargeIcon">
	/// 	True to show the large icon, and false to show the standard icon.
	/// </param>
	public static void CreateTogglePopUpButtonDefinition(RibbonPanel ribbonPanel, ObjectCollection collection, ButtonDefinition defaultButton, bool useLargeIcon)
	{
		try
		{
			_ = ribbonPanel.CommandControls.AddTogglePopup(defaultButton, collection, useLargeIcon, true, "", false);
		}
		catch (Exception ex)
		{
			System.Diagnostics.Debug.WriteLine($"Error in {nameof(CreateTogglePopUpButtonDefinition)}: {ribbonPanel.DisplayName}, {defaultButton.DisplayName}");
			System.Diagnostics.Debug.WriteLine(ex.Message);
		}
	}

	/// <summary>
	/// 	Creates a stack of buttons that changes the button icon when the user selects a tool.
	/// 	See the 'Display Mode' drop-down (Shaded, Hidden Edge, Wireframe) on the 'View' tab in parts and assemblies.
	/// </summary>
	/// <param name="ribbonPanel">
	/// 	The panel that gets the stack.
	/// </param>
	/// <param name="collection">
	/// 	The buttons in the stack.
	/// </param>
	/// <param name="useLargeIcon">
	/// 	True to show the large icon, and false to show the standard icon.
	/// </param>
	public static void CreateButtonPopUpButtonDefinition(RibbonPanel ribbonPanel, ObjectCollection collection, bool useLargeIcon)
	{
		try
		{
			_ = ribbonPanel.CommandControls.AddButtonPopup(collection, useLargeIcon, true, "", false);
		}
		catch (Exception ex)
		{
			System.Diagnostics.Debug.WriteLine($"Error in {nameof(CreateButtonPopUpButtonDefinition)}: {ribbonPanel.DisplayName}");
			System.Diagnostics.Debug.WriteLine(ex.Message);
		}
	}

	/// <summary>
	/// 	Creates a stack of buttons that shows the most recently used button.
	/// 	See the 'Circle' drop-down in the Sketch tab.
	/// </summary>
	/// <param name="ribbonPanel">
	/// 	The panel that gets the stack.
	/// </param>
	/// <param name="collection">
	/// 	The buttons in the stack.
	/// </param>
	/// <param name="useLargeIcon">
	/// 	True to show the large icon, and false to show the standard icon.
	/// </param>
	public static void CreateSplitMostRecentlyUsedButtonDefinition(RibbonPanel ribbonPanel, ObjectCollection collection, bool useLargeIcon)
	{
		try
		{
			_ = ribbonPanel.CommandControls.AddSplitButtonMRU(collection, useLargeIcon, true, "", false);
		}
		catch (Exception ex)
		{
			System.Diagnostics.Debug.WriteLine($"Error in {nameof(CreateSplitMostRecentlyUsedButtonDefinition)}: {ribbonPanel.DisplayName}");
			System.Diagnostics.Debug.WriteLine(ex.Message);
		}
	}
}
