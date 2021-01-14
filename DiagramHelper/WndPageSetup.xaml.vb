Public Class WndPageSetup


    Public Property PageWidth As Double
        Get
            Return GetValue(PageWidthProperty)
        End Get

        Set(ByVal value As Double)
            SetValue(PageWidthProperty, value)
        End Set
    End Property

    Public Shared ReadOnly PageWidthProperty As DependencyProperty = _
                           DependencyProperty.Register("PageWidth", _
                           GetType(Double), GetType(WndPageSetup), _
                           New PropertyMetadata(21.0))

    Public Property PageHeight As Double
        Get
            Return GetValue(PageHeightProperty)
        End Get

        Set(ByVal value As Double)
            SetValue(PageHeightProperty, value)
        End Set
    End Property

    Public Shared ReadOnly PageHeightProperty As DependencyProperty = _
                           DependencyProperty.Register("PageHeight", _
                           GetType(Double), GetType(WndPageSetup), _
                           New PropertyMetadata(29.7))

    Private Sub BtnOk_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = True
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
    End Sub


    Private Sub WndPageSetup_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        NumWidth.Focus()
        NumWidth.TextBox.SelectAll()
    End Sub
End Class
