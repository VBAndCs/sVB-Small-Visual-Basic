Imports System.Windows.Controls.Primitives

Public Class ColorBox
    Inherits ContentControl

    Shared WithEvents Pup As Popup
    Shared WithEvents Pkr As ColorPicker
    Dim WithEvents Box As ToggleButton

    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(ColorBox), New FrameworkPropertyMetadata(GetType(ColorBox)))
    End Sub

    Public Property Brush() As Brush
        Get
            Return CType(GetValue(BrushProperty), Brush)
        End Get
        Set(ByVal value As Brush)
            SetValue(BrushProperty, value)
        End Set
    End Property

    Public Shared ReadOnly BrushProperty As DependencyProperty = DependencyProperty.Register("Brush", GetType(Brush), GetType(ColorBox), New FrameworkPropertyMetadata(Nothing, AddressOf OnBrushChanged))

    Public Event BrushChanged(OldBrush As Brush, NewBrush As Brush)

    Sub OnBrushChanged(OldBrush As Brush, NewBrush As Brush)
        If Pup IsNot Nothing AndAlso Pup.IsOpen Then Return
        RaiseEvent BrushChanged(If(OldBrush Is Nothing, Nothing, OldBrush.CloneCurrentValue),
                                                   If(NewBrush Is Nothing, Nothing, NewBrush.CloneCurrentValue))
    End Sub

    Shared Sub OnBrushChanged(Cb As ColorBox, e As System.Windows.DependencyPropertyChangedEventArgs)
        Cb.OnBrushChanged(e.OldValue, e.NewValue)
    End Sub

    Public Property Color() As Color
        Get
            Return CType(GetValue(ColorProperty), Color)
        End Get
        Set(ByVal value As Color)
            SetValue(ColorProperty, value)
        End Set
    End Property

    Public Shared ReadOnly ColorProperty As DependencyProperty =
        DependencyProperty.Register(
            "Color",
            GetType(Color), GetType(ColorBox),
            New UIPropertyMetadata(Colors.Black, AddressOf OnColorChanged)
        )

    Public Event ColorChanged(OldColor As Color, NewColor As Color)
    Public Event Opening()
    Public Event Closed(Canceled As Boolean)

    Sub OnColorChanged(OldColor As Color, NewColor As Color)
        RaiseEvent ColorChanged(OldColor, NewColor)
    End Sub

    Shared Sub OnColorChanged(Cb As ColorBox, e As System.Windows.DependencyPropertyChangedEventArgs)
        Cb.OnColorChanged(e.OldValue, e.NewValue)
    End Sub

    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        Box = GetTemplateChild("PART_Box")
    End Sub

    Dim TmpBrush As Brush
    Private Sub Box_Checked(sender As Object, e As RoutedEventArgs) Handles Box.Checked
        RaiseEvent Opening()

        If ColorBox.Pup Is Nothing Then
            Pup = New Popup With {
                .StaysOpen = False,
                .AllowsTransparency = True,
                .HorizontalOffset = -1,
                .MinWidth = 290,
                .MaxHeight = 320,
                .Placement = PlacementMode.Bottom
            }
            Pkr = New ColorPicker With {
                .ShowPopUpColorDegrees = False,
                .MinWidth = Pup.MinWidth
            }
            Pup.Child = Pkr
        End If

        Pup.IsOpen = False
        Pup.Tag = Me

        If PreviewBox IsNot Nothing Then
            TmpBrush = PreviewBox.Background
            PreviewBox.Background = Me.Brush
        End If

        ColorBox.Pup.PlacementTarget = Box

        Try
            Pkr.Brush = Nothing ' To reset Popup pos
            ColorBox.Pup.IsOpen = True
            ' Set the Color after Popup appears to force it change it's Height
            Pkr.Brush = Brushes.Transparent
            Pkr.Brush = Me.Brush
            ColorBox.Pkr.Focus()

        Catch ex As Exception

        End Try

    End Sub

    Private Shared Sub Pup_Closed(sender As Object, e As EventArgs) Handles Pup.Closed
        Dim Cb As ColorBox = ColorBox.Pup.Tag
        If Not Scaped Then
            Cb.Brush = ColorBox.Pkr.Brush
            Cb.Color = ColorBox.Pkr.Color
            Cb.RaiseClosedEvent(False)
        Else
            Scaped = False
            Cb.RaiseClosedEvent(True)
        End If

        Dim Ui = Cb.Box.InputHitTest(Mouse.GetPosition(Cb.Box))
        If Ui Is Nothing Then Cb.Box.IsChecked = False
    End Sub

    Shared Scaped As Boolean = False
    Private Shared Sub Pup_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Pup.PreviewKeyDown
        Dim Cb As ColorBox = ColorBox.Pup.Tag
        If e.Key = Key.Escape Then
            ColorBox.Pup.IsOpen = False
            Cb.Box.IsChecked = False
            Scaped = True
            Cb.PreviewBox.Background = Cb.TmpBrush
            e.Handled = True
        ElseIf e.Key = Key.Enter Then
            ColorBox.Pup.IsOpen = False
            Cb.Box.IsChecked = False
            e.Handled = True
        End If
    End Sub

    Public Property PreviewBox As ContentControl
        Get
            Return GetValue(PreviewBoxProperty)
        End Get

        Set(ByVal value As ContentControl)
            SetValue(PreviewBoxProperty, value)
        End Set
    End Property

    Public Shared ReadOnly PreviewBoxProperty As DependencyProperty =
              DependencyProperty.Register("PreviewBox",
              GetType(Control), GetType(ColorBox),
              New PropertyMetadata(Nothing)
        )

    Private Shared Sub Pkr_ColorChanged(sender As Object, e As ColorChangedEventArgs) Handles Pkr.ColorChanged
        Dim Cb As ColorBox = ColorBox.Pup.Tag
        If Cb Is Nothing Then Return
        If Cb.PreviewBox IsNot Nothing Then Cb.PreviewBox.Background = Pkr.Brush 'Is Nothing, Nothing, Pkr.Brush.CloneCurrentValue)
    End Sub

    Private Sub RaiseClosedEvent(Canceled As Boolean)
        RaiseEvent Closed(Canceled)
    End Sub

End Class
