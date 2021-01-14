Friend Class Arrow

    Public Property Angle As Double
        Get
            Return GetValue(AngleProperty)
        End Get

        Set(ByVal value As Double)
            SetValue(AngleProperty, value)
            Dim Rt = CType(Me.LayoutTransform, RotateTransform)
            Rt.Angle = value
        End Set
    End Property

    Public Shared ReadOnly AngleProperty As DependencyProperty = _
                           DependencyProperty.Register("Angle", _
                           GetType(Double), GetType(Arrow), _
                           New PropertyMetadata(0.0))

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.RenderTransformOrigin = New Point(0.5, 0.5)
        Me.LayoutTransform = New RotateTransform(0)
    End Sub
End Class
