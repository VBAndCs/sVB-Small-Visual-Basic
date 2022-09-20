Friend Class Arrow

    Public Property Angle As Double
        Get
            Return GetValue(AngleProperty)
        End Get

        Set(value As Double)
            SetValue(AngleProperty, value)
            Dim rt = CType(Me.LayoutTransform, RotateTransform)
            rt.Angle = value
        End Set
    End Property

    Public Shared ReadOnly AngleProperty As DependencyProperty = _
                           DependencyProperty.Register("Angle", _
                           GetType(Double), GetType(Arrow), _
                           New PropertyMetadata(0.0))

    Public Sub New(angle As Integer)
        Me.New()
        Me.Angle = angle
    End Sub

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.RenderTransformOrigin = New Point(0.5, 0.5)
        Me.LayoutTransform = New RotateTransform(0)
    End Sub
End Class
