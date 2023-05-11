Public Class FontDialog

    Private Shared _FontProperties As List(Of DependencyProperty)
    Public Shared ReadOnly Property FontProperties() As List(Of DependencyProperty)
        Get
            If _FontProperties Is Nothing Then _FontProperties = FontChooser.GetFontProperties
            Return _FontProperties
        End Get
    End Property

    Public Shared Property ShowSymbolFonts As Boolean = False

    Private Shared FontChooserWref As WeakReference

    Public Shared Function Show(TargetControl As DependencyObject, Optional SampleText As String = "") As Boolean
        Dim FontPkr As New WndFont()
        Dim _FontChooser As FontChooser
        If FontChooserWref Is Nothing OrElse Not FontChooserWref.IsAlive Then
            _FontChooser = New FontChooser
            FontChooserWref = New WeakReference(_FontChooser)
        Else
            _FontChooser = FontChooserWref.Target
        End If

        FontPkr.Content = _FontChooser

        _FontChooser.SetPropertiesFromObject(TargetControl)
        _FontChooser.PreviewSampleText = SampleText

        If FontPkr.ShowDialog().Value Then
            _FontChooser.ApplyPropertiesToObject(TargetControl)
            Show = True
        End If

        FontPkr.Content = Nothing
        Return Show
    End Function
End Class
