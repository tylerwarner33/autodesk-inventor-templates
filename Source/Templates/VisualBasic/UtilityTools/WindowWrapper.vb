Module WindowWrapper

    ''' <summary>
    ''' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ''' This is primarily used for parenting a dialog to the Inventor window.
    ''' <code><example>
    '''     myForm.Show(New WindowWrapper(InventorApplication.MainFrameHWND))
    ''' </example></code>
    ''' </summary>
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Public Sub New(ByVal handle As IntPtr)
            _handle = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _handle
            End Get
        End Property

        Private _handle As IntPtr
    End Class

End Module
