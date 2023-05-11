Imports System.Text

Public Class TextBoxBehavior
    Public Shared Function GetSelectAllTextOnFocus(ByVal textBox As TextBox) As Boolean
        Return CBool(textBox.GetValue(SelectAllTextOnFocusProperty))
    End Function

    Public Shared Sub SetSelectAllTextOnFocus(ByVal textBox As TextBox, ByVal value As Boolean)
        textBox.SetValue(SelectAllTextOnFocusProperty, value)
    End Sub

    Public Shared ReadOnly SelectAllTextOnFocusProperty As DependencyProperty = DependencyProperty.RegisterAttached("SelectAllTextOnFocus", GetType(Boolean), GetType(TextBoxBehavior), New UIPropertyMetadata(False, AddressOf OnSelectAllTextOnFocusChanged))

    Private Shared Sub OnSelectAllTextOnFocusChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim textBox = TryCast(d, TextBox)
        If textBox Is Nothing Then
            Return
        End If

        If TypeOf e.NewValue Is Boolean = False Then
            Return
        End If

        If CBool(e.NewValue) Then
            AddHandler textBox.GotFocus, AddressOf SelectAll
            AddHandler textBox.PreviewMouseDown, AddressOf IgnoreMouseButton
        Else
            RemoveHandler textBox.GotFocus, AddressOf SelectAll
            RemoveHandler textBox.PreviewMouseDown, AddressOf IgnoreMouseButton
        End If
    End Sub

    Private Shared Sub SelectAll(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim textBox = TryCast(e.OriginalSource, TextBox)
        If textBox Is Nothing Then
            Return
        End If
        textBox.SelectAll()
    End Sub

    Private Shared Sub IgnoreMouseButton(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs)
        Dim textBox = TryCast(sender, TextBox)
        If textBox Is Nothing OrElse textBox.IsKeyboardFocusWithin Then
            Return
        End If

        e.Handled = True
        textBox.Focus()
    End Sub
End Class
