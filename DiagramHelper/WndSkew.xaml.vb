Public Class WndSkew

    Public Property SkewTransform As SkewTransform
        Get
            Return GetValue(SkewTransformProperty)
        End Get

        Set(value As SkewTransform)
            SetValue(SkewTransformProperty, value)
        End Set
    End Property

    Public Shared ReadOnly SkewTransformProperty As DependencyProperty = _
                           DependencyProperty.Register("SkewTransform", _
                           GetType(SkewTransform), GetType(WndSkew))

    Private Sub BtnOk_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = True
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
    End Sub

    Private Sub WndSkew_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        NumAngleX.Focus()
        NumAngleX.TextBox.SelectAll()
    End Sub

    Private Sub WndSkew_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If Me.SkewTransform Is Nothing Then Me.SkewTransform = New SkewTransform(0, 0, 0.5, 0.5)        
    End Sub
End Class
