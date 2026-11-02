namespace InventorAddIn.UtilityTools;

/// <summary>
/// 	Wraps a Win32 window handle as an <see cref="IWin32Window" />.
/// 	Use it to make the Inventor window the parent of a dialog.
/// </summary>
/// <example>
/// 	<code>
/// 		myForm.Show(new WindowWrapper(new IntPtr(Globals.InventorApplication.MainFrameHWND)));
/// 	</code>
/// </example>
/// <param name="handle">
/// 	The window handle.
/// </param>
internal sealed class WindowWrapper(IntPtr handle) : IWin32Window
{
	/// <inheritdoc />
	public IntPtr Handle { get; } = handle;
}
