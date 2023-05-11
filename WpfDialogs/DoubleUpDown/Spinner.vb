Imports System.Windows.Markup
Imports System.Windows.Controls.Primitives

Friend Class Spinner
    Inherits ContentControl

    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(Spinner), New FrameworkPropertyMetadata(GetType(Spinner)))
    End Sub

    Public Shared ReadOnly ValidSpinDirectionProperty As DependencyProperty = DependencyProperty.Register("ValidSpinDirection", GetType(ValidSpinDirections), GetType(Spinner), New PropertyMetadata(ValidSpinDirections.Increase Or ValidSpinDirections.Decrease, AddressOf OnValidSpinDirectionPropertyChanged))
    Public Property ValidSpinDirection() As ValidSpinDirections
        Get
            Return CType(GetValue(ValidSpinDirectionProperty), ValidSpinDirections)
        End Get
        Set(ByVal value As ValidSpinDirections)
            SetValue(ValidSpinDirectionProperty, value)
        End Set
    End Property

    Public Shared Sub OnValidSpinDirectionPropertyChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oldvalue As ValidSpinDirections = CType(e.OldValue, ValidSpinDirections)
        Dim newvalue As ValidSpinDirections = CType(e.NewValue, ValidSpinDirections)
    End Sub

    Public Event Spin As EventHandler(Of SpinEventArgs)
    Protected Overridable Sub OnSpin(ByVal e As SpinEventArgs)
        Dim handler As EventHandler(Of SpinEventArgs) = SpinEvent
        If handler IsNot Nothing Then
            handler(Me, e)
        End If
    End Sub

    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()

        Dim IncreaseButton As RepeatButton = GetTemplateChild("PART_IncreaseButton")
        AddHandler IncreaseButton.Click, AddressOf IncreaseButton_Click

        Dim DecreaseButton As RepeatButton = GetTemplateChild("PART_DecreaseButton")
        AddHandler DecreaseButton.Click, AddressOf DecreaseButton_Click

    End Sub

    Private Sub IncreaseButton_Click(sender As Object, e As RoutedEventArgs)
        RaiseEvent Spin(Me, New SpinEventArgs(SpinDirection.Increase))
    End Sub

    Private Sub DecreaseButton_Click(sender As Object, e As RoutedEventArgs)
        RaiseEvent Spin(Me, New SpinEventArgs(SpinDirection.Decrease))
    End Sub


    Protected Overrides Sub OnMouseWheel(ByVal e As MouseWheelEventArgs)
        MyBase.OnMouseWheel(e)
        Dim F As Control = Keyboard.FocusedElement
        If Not F.Parent Is Me Then Return
        If e.Delta < 0 Then
            RaiseEvent Spin(Me, New SpinEventArgs(SpinDirection.Decrease))
        ElseIf 0 < e.Delta Then
            RaiseEvent Spin(Me, New SpinEventArgs(SpinDirection.Increase))
        End If

        e.Handled = True
    End Sub

    Protected Overrides Sub OnPreviewKeyDown(ByVal e As KeyEventArgs)
        Select Case e.Key
            Case Key.Up
                RaiseEvent Spin(Me, New SpinEventArgs(SpinDirection.Increase))
                e.Handled = True
            Case Key.Down
                RaiseEvent Spin(Me, New SpinEventArgs(SpinDirection.Decrease))
                e.Handled = True
                Exit Select
        End Select
    End Sub

End Class