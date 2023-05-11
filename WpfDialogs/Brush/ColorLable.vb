Friend Class ColorLable
    Inherits Border

    Dim _IsSelected As Boolean
    Dim Brdr As New Border

    Dim _Color As Color

    Property Color As Color
        Get
            Return _Color
        End Get
        Set(value As Color)
            _Color = value
            Brdr.Background = New SolidColorBrush(value)
        End Set
    End Property

    Friend Property IsSelected As Boolean
        Get
            Return _IsSelected
        End Get
        Set(value As Boolean)
            If _IsSelected = value Then Return
            _IsSelected = value
            Me.BorderThickness = New Thickness(If(value, 2, 0.5))
        End Set
    End Property

    Public Sub New(Color As Color)
        MyBase.New()
        Me.BorderThickness = New Thickness(0.5)
        Me.BorderBrush = Brushes.Black
        Me.Width = 26
        Me.Height = 26
        Me.Margin = New Thickness(2)
        Me.Padding = New Thickness(2)
        Me.Child = Brdr
        _Color = Color
        Me.ToolTip = ColorHelper.GetColorName(Color)
        Brdr.Background = New SolidColorBrush(Color)
    End Sub

    Public Sub New(Hue As Double, Saturation As Double, Brightness As Double)
        Me.New(ColorHelper.ColorFromHSB(Hue, Saturation, Brightness))
    End Sub

End Class
