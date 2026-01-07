Imports Inventor
Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.VisualBasic.Compatibility.VB6

Friend MustInherit Class Button

#Region "Data Members"


    Private m_buttonDefinition As ButtonDefinition

#End Region

#Region "Properties"


    Public ReadOnly Property ButtonDefinition() As Inventor.ButtonDefinition
        Get
            ButtonDefinition = m_buttonDefinition
        End Get
    End Property

#End Region

#Region "Methods"

    Public Sub New(ByVal Environment As String, CustomDrawingTab As RibbonTab,
                   RibbonPanel As RibbonPanel, useLargeIcon As Boolean, isInButtonStack As Boolean)

        Try
            'get the images to use for the button
            Dim largeIcon As IPictureDisp = Nothing
            Dim standardIcon As IPictureDisp = Nothing
            Dim toolTipImage As IPictureDisp = Nothing

            'this is the text the user sees on the button
            Dim buttonLabel As String = Nothing

            'text that displays when the user hovers over the button
            Dim toolTip_Simple As String = Nothing
            Dim toolTip_Expanded As String = Nothing
            Dim useProgressToolTip As Boolean = True


            'create button definition
            m_buttonDefinition = CreateButtonDefintion.CreateButtonDef(Environment, CustomDrawingTab, RibbonPanel, useLargeIcon,
                                                            isInButtonStack, useProgressToolTip,
                                                            buttonLabel, toolTip_Simple, toolTip_Expanded,
                                                            standardIcon, largeIcon, toolTipImage)
            'enable the button
            m_buttonDefinition.Enabled = True

            'connect the button event sink
            AddHandler m_buttonDefinition.OnExecute, AddressOf Me.ButtonDefinition_OnExecute

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Protected MustOverride Sub ButtonDefinition_OnExecute(ByVal context As NameValueMap)

#End Region

End Class
