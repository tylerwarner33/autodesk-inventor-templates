namespace InventorAddIn.UtilityTools;

/// <summary>
/// 	Converts a .NET image to the <see cref="stdole.IPictureDisp" /> type that the Inventor API requires.
/// </summary>
/// <remarks>
/// 	<c>GetIPictureDispFromPicture</c> is a protected static member of <see cref="AxHost" />, so only a subclass can call it.
/// </remarks>
internal sealed class PictureDispConverter : AxHost
{
	private PictureDispConverter()
		: base(string.Empty)
	{
	}

	/// <summary>
	/// 	Converts an image.
	/// </summary>
	/// <param name="image">
	/// 	The image, ex. a bitmap from the project resources.
	/// </param>
	/// <returns>
	/// 	The picture that Inventor takes.
	/// </returns>
	public static stdole.IPictureDisp ToIPictureDisp(System.Drawing.Image image)
	{
		return GetIPictureDispFromPicture(image) as stdole.IPictureDisp
			?? throw new InvalidOperationException("The image could not be converted to a picture.");
	}
}
