
Public Class OpenedPagesExplorer
    Inherits Explorer

    Protected Overrides ReadOnly Property ItemsSource As Specialized.INotifyCollectionChanged
        Get
            Return Designer.FormNames
        End Get
    End Property

    Protected Overrides Sub OnSelectionChanged()
        Designer.SwitchTo(Designer.FormKeys(FilesList.SelectedIndex))
    End Sub

    Protected Overrides Sub OnDeleteItem()
        If Designer.IsNew AndAlso Designer.Pages.Count = 1 OrElse FilesList.SelectedIndex = -1 Then
            Beep()
            Return
        End If

        Designer.ClosePage()
    End Sub

    Protected Overrides Function OnCommit(newName As String) As Boolean
        If Helper.FormNameExists(Designer, newName) Then
            Return False
        End If

        Return Designer.ChangeFormName(newName)
    End Function

    Dim firstTime As Boolean = True

    Private Sub Explorer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If firstTime Then
            firstTime = False
            FilesList.ItemsSource = ItemsSource
            If FilesList.Items.Count > 0 Then FilesList.SelectedIndex = 0
        End If
    End Sub

    Public Property Designer As Designer
        Get
            Return GetValue(DesignerProperty)
        End Get

        Set(value As Designer)
            SetValue(DesignerProperty, value)
        End Set
    End Property

    Public Shared ReadOnly DesignerProperty As DependencyProperty =
               DependencyProperty.Register("Designer",
               GetType(Designer), GetType(OpenedPagesExplorer))

End Class
