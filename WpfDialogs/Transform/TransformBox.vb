Imports System.Windows.Controls.Primitives

Public Class TransformBox
    Inherits ContentControl

    Shared WithEvents Pup As Popup
    Shared WithEvents Pkr As TransformPicker
    Dim WithEvents Box As ToggleButton

    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(TransformBox), New FrameworkPropertyMetadata(GetType(TransformBox)))
    End Sub

    Public Property Transform() As Transform
        Get
            Return CType(GetValue(TransformProperty), Transform)
        End Get
        Set(ByVal value As Transform)
            SetValue(TransformProperty, value)
        End Set
    End Property

    Public Shared ReadOnly TransformProperty As DependencyProperty = DependencyProperty.Register("Transform", GetType(Transform), GetType(TransformBox))

    Public Event TransformChanged(TransBox As TransformBox, OldTransform As Transform, NewTransform As Transform)

    Sub OnTransformChanged(TransBox As TransformBox, OldTransform As Transform, NewTransform As Transform)
        RaiseEvent TransformChanged(TransBox, OldTransform, NewTransform)
    End Sub

    Public Property TargetWidth As Double
        Get
            Return GetValue(TargetWidthProperty)
        End Get

        Set(ByVal value As Double)
            SetValue(TargetWidthProperty, value)
        End Set
    End Property

    Public Shared ReadOnly TargetWidthProperty As DependencyProperty = _
                       DependencyProperty.Register("TargetWidth", _
                       GetType(Double), GetType(TransformBox), _
                       New PropertyMetadata(1.0))


    Public Property TargetHeight As Double
        Get
            Return GetValue(TargetHeightProperty)
        End Get

        Set(ByVal value As Double)
            SetValue(TargetHeightProperty, value)
        End Set
    End Property

    Public Shared ReadOnly TargetHeightProperty As DependencyProperty = _
                           DependencyProperty.Register("TargetHeight", _
                           GetType(Double), GetType(TransformBox), _
                           New PropertyMetadata(1.0))

    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        Box = GetTemplateChild("PART_Box")
    End Sub

    Private Sub Box_Checked(sender As Object, e As RoutedEventArgs) Handles Box.Checked
        If TransformBox.Pup Is Nothing Then
            Pup = New Popup With {.StaysOpen = False, .AllowsTransparency = True, .HorizontalOffset = -1,
                                  .MinWidth = 290, .MaxHeight = 360, .Placement = PlacementMode.Bottom}
            Pkr = New TransformPicker
            Pup.Child = Pkr

        End If

        TransformBox.Pup.IsOpen = False
        Pup.Tag = Me

        TransformBox.Pup.PlacementTarget = Box

        Try
            TransformBox.Pup.IsOpen = True
            If Pkr.Transform IsNot Me.Transform Then Pkr.Transform = Me.Transform
            Pkr.TargetWidth = Me.TargetWidth
            Pkr.TargetHeight = Me.TargetHeight
            TransformBox.Pkr.Focus()
        Catch ex As Exception

        End Try

    End Sub

    Private Shared Sub Pup_Closed(sender As Object, e As EventArgs) Handles Pup.Closed
        Dim Cb As TransformBox = TransformBox.Pup.Tag
        If Not Scaped Then
            Cb.Transform = TransformBox.Pkr.Transform
        Else
            Scaped = False
        End If
        Dim Ui = Cb.Box.InputHitTest(Mouse.GetPosition(Cb.Box))
        If Ui Is Nothing Then Cb.Box.IsChecked = False
    End Sub

    Shared Scaped As Boolean = False
    Private Shared Sub Pup_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Pup.PreviewKeyDown
        Dim Cb As TransformBox = TransformBox.Pup.Tag
        If e.Key = Key.Escape Then
            TransformBox.Pup.IsOpen = False
            Cb.Box.IsChecked = False
            Scaped = True
            e.Handled = True
        ElseIf e.Key = Key.Enter Then
            TransformBox.Pup.IsOpen = False
            Cb.Box.IsChecked = False
            e.Handled = True
        End If
    End Sub

    Private Shared Sub Pkr_TransformChanged(Tp As TransformPicker, OldTransform As Transform, NewTransform As Transform) Handles Pkr.TransformChanged
        Dim TransformBox As TransformBox = Pup.Tag
        If Tp Is Nothing Then Return
        TransformBox.OnTransformChanged(TransformBox, OldTransform, NewTransform)
    End Sub
End Class
