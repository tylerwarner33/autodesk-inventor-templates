Imports System.Runtime.InteropServices

Module PictureConverter

    ''' <summary>
    ''' Class used to convert bitmaps and icons from their .NET native types into an <see cref="stdole.IPictureDisp"/> object which is what the Inventor API requires.
    ''' A typical usage is shown below where MyIcon is a bitmap or icon that's available as a resource of the project.
    ''' </summary>
    Public NotInheritable Class PictureDispConverter

        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)>
        Private Shared Function OleCreatePictureIndirect(
            ByRef picdesc As PICTDESC_BITMAP,
            ByRef iid As Guid,
            <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
        End Function

        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)>
        Private Shared Function OleCreatePictureIndirect(
            ByRef picdesc As PICTDESC_ICON,
            ByRef iid As Guid,
            <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
        End Function

        Shared iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

        ' Picture Types
        Private Const PICTYPE_BITMAP As Short = 1
        Private Const PICTYPE_ICON As Short = 3

        <StructLayout(LayoutKind.Sequential)>
        Private Structure PICTDESC_BITMAP
            Public cbSizeOfStruct As Integer
            Public picType As Integer
            Public hbitmap As IntPtr
            Public hpal As IntPtr

            Public Sub New(bmp As System.Drawing.Bitmap)
                cbSizeOfStruct = Marshal.SizeOf(GetType(PICTDESC_BITMAP))
                picType = PICTYPE_BITMAP
                hbitmap = bmp.GetHbitmap()
                hpal = IntPtr.Zero
            End Sub
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure PICTDESC_ICON
            Public cbSizeOfStruct As Integer
            Public picType As Integer
            Public hicon As IntPtr

            Public Sub New(ico As System.Drawing.Icon)
                cbSizeOfStruct = Marshal.SizeOf(GetType(PICTDESC_ICON))
                picType = PICTYPE_ICON
                hicon = ico.Handle
            End Sub
        End Structure

        Public Shared Function ToIPictureDisp(bmp As System.Drawing.Bitmap) As stdole.IPictureDisp
            Dim desc As New PICTDESC_BITMAP(bmp)
            Return OleCreatePictureIndirect(desc, iPictureDispGuid, False)
        End Function

        Public Shared Function ToIPictureDisp(ico As System.Drawing.Icon) As stdole.IPictureDisp
            Dim desc As New PICTDESC_ICON(ico)
            Return OleCreatePictureIndirect(desc, iPictureDispGuid, False)
        End Function

    End Class

End Module
