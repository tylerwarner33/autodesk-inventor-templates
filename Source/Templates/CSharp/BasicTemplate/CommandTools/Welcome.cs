using System.Reflection;

namespace InventorAddIn.CommandTools;

/// <summary>
/// 	This command has no button. It runs on an event.
/// </summary>
internal static class Welcome
{
	/// <summary>
	/// 	Shows a message.
	/// </summary>
	/// <param name="documentName">
	/// 	The name of the document that started the event.
	/// </param>
	/// <param name="eventName">
	/// 	The name of the event.
	/// </param>
	public static void AddinHello(string documentName, string eventName)
	{
		if (documentName.Length == 0)
			documentName = "a new file";

		string? addInName = Assembly.GetExecutingAssembly().GetName().Name;

		_ = MessageBox.Show(
			$"My Name is {addInName}, I'm an add-in.\nI fired on the {eventName.Replace("k", "")} event of:\n{documentName}",
			$"Hello {System.Environment.UserName}");
	}
}
