namespace InventorAddIn;

/// <summary>
/// 	Holds the values that all of the add-in code uses.
/// </summary>
internal static class Globals
{
	/// <summary>
	/// 	The unique ID of this add-in, without braces.
	/// </summary>
	public const string SimpleAddInClientId = "CLIENT_GUID_PLACEHOLDER";

	/// <summary>
	/// 	The unique ID of this add-in, with braces.
	/// </summary>
	public const string AddInClientId = "{" + SimpleAddInClientId + "}";

	/// <summary>
	/// 	Gets or sets the Inventor application that loaded this add-in.
	/// </summary>
	public static Inventor.Application InventorApplication { get; set; } = null!;
}
