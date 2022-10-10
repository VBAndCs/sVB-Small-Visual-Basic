
Imports System.Collections
Imports System.Windows
Imports DiagramHelper
Imports Microsoft.SmallVisualBasic.Shell

Namespace Microsoft.SmallVisualBasic.Documents

    Public Class OpenedDocsExplorer
        Inherits Explorer

        Protected Overrides ReadOnly Property ItemsSource As Specialized.INotifyCollectionChanged
            Get
                Return MdiViews.Items
            End Get
        End Property

        Protected Overrides Sub OnSelectionChanged()
            MdiViews.SelectedItem = FilesList.SelectedItem
        End Sub

        Protected Overrides Sub OnDeleteItem()
            Dim view = CType(FilesList.SelectedItem, MdiView)
            Helper.MainWindow.CloseView(view)
        End Sub

        Protected Overrides Function OnBeginEdit() As Boolean
            Return False
        End Function

        Protected Overrides Function OnCommit(newName As String) As Boolean
            Return True
        End Function

        Dim firstTime As Boolean = True

        Private Sub Explorer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
            If firstTime Then
                firstTime = False
                FilesList.ItemsSource = ItemsSource
                FilesList.DisplayMemberPath = "Document.Title"
                If FilesList.Items.Count > 0 Then FilesList.SelectedIndex = 0
            End If
        End Sub

        Public Property MdiViews As MdiViewsControl
            Get
                Return GetValue(MdiViewsProperty)
            End Get

            Set(value As MdiViewsControl)
                SetValue(MdiViewsProperty, value)
            End Set
        End Property

        Public Shared ReadOnly MdiViewsProperty As DependencyProperty =
               DependencyProperty.Register("MdiViews",
               GetType(MdiViewsControl), GetType(OpenedDocsExplorer))

    End Class

End Namespace
