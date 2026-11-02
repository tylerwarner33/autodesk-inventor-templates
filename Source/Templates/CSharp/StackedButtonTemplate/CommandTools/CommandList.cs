using Inventor;
using InventorAddIn.Properties;
using InventorAddIn.UtilityTools;
using System.Diagnostics;

namespace InventorAddIn.CommandTools;

/// <summary>
/// 	Writes all of the Inventor commands to a text file.
/// 	<see cref="CreateButton" /> defines the label, the icons, and the tool tip of the button.
/// </summary>
internal static class CommandList
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
		stdole.IPictureDisp largeIcon = PictureDispConverter.ToIPictureDisp(Resources.Snake_32);
		stdole.IPictureDisp standardIcon = PictureDispConverter.ToIPictureDisp(Resources.Snake_16);
		stdole.IPictureDisp? toolTipImage = null;

		// This is the text the user sees on the button.
		string buttonLabel = "Command\nList";

		// Text that displays when the user hovers over the button.
		string toolTipSimple = "Writes out all Inventor Commands to a text file";
		string? toolTipExpanded = null;

		// Set to true to use a progressive tool tip, and false to use a simple tool tip.
		bool useProgressToolTip = true;

		// Only used if 'useProgressToolTip' is true.
		if (useProgressToolTip)
		{
			toolTipImage = PictureDispConverter.ToIPictureDisp(Resources.AU_ToolTip);

			toolTipExpanded =
				"\u2022 This tool creates a text file and writes all the command information out\n" +
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
	/// 	Writes the commands to a text file and opens the file.
	/// </summary>
	public static void RunCommandCode()
	{
		string folder = @"C:\temp\";
		string fileName = $"{folder}Inventor Command List.txt";

		_ = System.IO.Directory.CreateDirectory(folder);

		using (System.IO.StreamWriter streamWriter = new(fileName))
		{
			streamWriter.WriteLine("Inventor Command List");
			streamWriter.WriteLine();
			streamWriter.WriteLine("Use Example:");
			streamWriter.WriteLine("1) Find the command in the list such as:");
			streamWriter.WriteLine("             AppViewCubeHomeCmd ");
			streamWriter.WriteLine("2) Then add it to the command manager line such as this, and use this line in your code:");
			streamWriter.WriteLine("            ThisApplication.CommandManager.ControlDefinitions.Item(\"AppViewCubeHomeCmd\").Execute");
			streamWriter.WriteLine("____________________________________________________________________________________");
			streamWriter.WriteLine();

			foreach (ControlDefinition controlDefinition in Globals.InventorApplication.CommandManager.ControlDefinitions)
			{
				if (controlDefinition.DescriptionText.Length == 0)
					streamWriter.WriteLine(controlDefinition.InternalName);
				else
					streamWriter.WriteLine($"{controlDefinition.InternalName}\t\t\t\t\t\t  \"{controlDefinition.DescriptionText}\"");
			}
		}

		_ = Process.Start(new ProcessStartInfo(fileName) { UseShellExecute = true });
	}
}
