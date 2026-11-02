Imports Inventor
Imports System.Drawing
Imports System.Windows.Forms

Friend MustInherit Class Button

    Private _buttonDefinition As ButtonDefinition

    Public ReadOnly Property ButtonDefinition() As Inventor.ButtonDefinition
        Get
            ButtonDefinition = _buttonDefinition
        End Get
    End Property

    Public Sub New(
        ByVal Environment As String,
        CustomDrawingTab As RibbonTab,
        RibbonPanel As RibbonPanel,
        useLargeIcon As Boolean,
        isInButtonStack As Boolean)

        Try
            ' Get the images to use for the button.
            Dim largeIcon As IPictureDisp = Nothing
            Dim standardIcon As IPictureDisp = Nothing
            Dim toolTipImage As IPictureDisp = Nothing

            ' This is the text the user sees on the button.
            Dim buttonLabel As String = Nothing

            ' Text that displays when the user hovers over the button.
            Dim toolTip_Simple As String = Nothing
            Dim toolTip_Expanded As String = Nothing
            Dim useProgressToolTip As Boolean = True

            ' Create button definition.
            _buttonDefinition = CreateButtonDefintion.CreateButtonDef(
                Environment,
                CustomDrawingTab,
                RibbonPanel,
                useLargeIcon,
                isInButtonStack,
                useProgressToolTip,
                buttonLabel,
                toolTip_Simple,
                toolTip_Expanded,
                standardIcon,
                largeIcon,
                toolTipImage)

            ' Enable the button.
            _buttonDefinition.Enabled = True

            ' Connect the button event sink.
            AddHandler _buttonDefinition.OnExecute, AddressOf Me.ButtonDefinition_OnExecute

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Protected MustOverride Sub ButtonDefinition_OnExecute(ByVal context As NameValueMap)

End Class
