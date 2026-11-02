Imports System.Runtime.InteropServices

Module PictureConverter

    ''' <summary>
    ''' Class used to convert bitmaps and icons from their .NET native types into an <see cref="stdole.IPictureDisp"/> object which is what the Inventor API requires.
    ''' A typical usage is shown below where MyIcon is a bitmap or icon that's available as a resource of the project.
    ''' </summary>
    Public NotInheritable Class PictureDispConverter

        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)>
        Private Shared Function OleCreatePictureIndirect(
            ByRef pictureDescription As PictureDescriptionBitmap,
            ByRef interfaceId As Guid,
            <MarshalAs(UnmanagedType.Bool)> ByVal ownsHandle As Boolean) As stdole.IPictureDisp
        End Function

        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)>
        Private Shared Function OleCreatePictureIndirect(
            ByRef pictureDescription As PictureDescriptionIcon,
            ByRef interfaceId As Guid,
            <MarshalAs(UnmanagedType.Bool)> ByVal ownsHandle As Boolean) As stdole.IPictureDisp
        End Function

        Shared _iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

        ' Picture Types
        Private Const PictureTypeBitmap As Short = 1
        Private Const PictureTypeIcon As Short = 3

        <StructLayout(LayoutKind.Sequential)>
        Private Structure PictureDescriptionBitmap
            Public cbSizeOfStruct As Integer
            Public picType As Integer
            Public hbitmap As IntPtr
            Public hpal As IntPtr

            Public Sub New(bitmap As System.Drawing.Bitmap)
                cbSizeOfStruct = Marshal.SizeOf(GetType(PictureDescriptionBitmap))
                picType = PictureTypeBitmap
                hbitmap = bitmap.GetHbitmap()
                hpal = IntPtr.Zero
            End Sub
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure PictureDescriptionIcon
            Public cbSizeOfStruct As Integer
            Public picType As Integer
            Public hicon As IntPtr

            Public Sub New(icon As System.Drawing.Icon)
                cbSizeOfStruct = Marshal.SizeOf(GetType(PictureDescriptionIcon))
                picType = PictureTypeIcon
                hicon = icon.Handle
            End Sub
        End Structure

        Public Shared Function ToIPictureDisp(bitmap As System.Drawing.Bitmap) As stdole.IPictureDisp
            Dim description As New PictureDescriptionBitmap(bitmap)
            Return OleCreatePictureIndirect(description, _iPictureDispGuid, False)
        End Function

        Public Shared Function ToIPictureDisp(icon As System.Drawing.Icon) As stdole.IPictureDisp
            Dim description As New PictureDescriptionIcon(icon)
            Return OleCreatePictureIndirect(description, _iPictureDispGuid, False)
        End Function

    End Class

End Module
