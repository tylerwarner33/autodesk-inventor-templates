Module WindowWrapper

    ''' <summary>
    ''' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ''' This is primarily used for parenting a dialog to the Inventor window.
    ''' <code><example>
    '''     myForm.Show(New WindowWrapper(_inventorApplication.MainFrameHWND))
    ''' </example></code>
    ''' </summary>
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

        Private _hwnd As IntPtr
    End Class

End Module
