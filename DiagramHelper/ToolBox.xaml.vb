Public Class ToolBox

    Dim tabsLoaded As Boolean
    Dim Items As New List(Of ToolBoxItem)

    Private Sub ToolBox_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If tabsLoaded Then Return
        tabsLoaded = True

        Try
            Dim appDir = IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()(0))
            Dim toolBoxDir = IO.Path.Combine(appDir, "ToolBox")
            For Each FolderName In IO.Directory.GetDirectories(toolBoxDir)
                AddTab(FolderName)
            Next
        Catch
        End Try
    End Sub

    Dim IsFirstTab As Boolean = True

    Sub AddTab(folderName As String)
        Dim Expan As New Expander With {.IsExpanded = IsFirstTab}
        Expan.Style = Me.FindResource("ExpanderStyle")
        IsFirstTab = False
        Expan.Header = IO.Path.GetFileNameWithoutExtension(folderName)

        Dim WrpPnl As New WrapPanel With {
            .Background = Brushes.White,
            .Margin = New Thickness(10, 0, 0, 0)
        }
        Dim Scv As New ScrollViewer With {
            .VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            .MaxHeight = 120,
            .Margin = New Thickness(1),
            .Background = Brushes.White,
            .Content = WrpPnl
        }
        Expan.Content = Scv
        Dim Files = IO.Directory.GetFiles(folderName, "*.xaml")
        Array.Sort(Files)
        For Each FileName In Files
            If IO.Path.GetFileNameWithoutExtension(FileName).ToLower.EndsWith("style") Then
                Dim WrpPnl2 = Helper.CreateWrapPanel(FileName)
                Do
                    Try
                        Dim Diagram As UIElement = WrpPnl2.Children(0)
                        WrpPnl2.Children.Remove(Diagram)
                        Dim Item As New ToolBoxItem(Diagram)
                        WrpPnl.Children.Add(Item)
                        Item.ToolBox = Me
                        Items.Add(Item)
                    Catch ex As Exception

                    End Try
                Loop While WrpPnl2.Children.Count > 0

            Else
                Try
                    Dim diagram = Helper.CreateDiagram(FileName)
                    Dim Item As New ToolBoxItem(diagram, FileName)
                    WrpPnl.Children.Add(Item)
                    Item.ToolBox = Me
                    Items.Add(Item)
                Catch ex As Exception
                End Try
            End If
        Next
        ToolBoxTabs.Children.Add(Expan)
    End Sub


    Dim _selectedItem As ToolBoxItem
    Friend Property SelectedItem As ToolBoxItem
        Get
            Return _selectedItem
        End Get

        Set(value As ToolBoxItem)
            _selectedItem = value
            Designer.SelectedToolBoxItem = value
            If value Is Nothing Then
                Designer.Cursor = Nothing
            Else
                Designer.Cursor = Cursors.Pen
            End If
        End Set
    End Property

    Public Sub DeSelectAll()
        For Each item In Items
            item.IsSelected = False
        Next
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
                           GetType(Designer), GetType(ToolBox))

    Private Sub UserControl_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Escape Then
            Me.DeSelectAll()
        End If
    End Sub

    Private Sub ToolBoxTabs_PreviewMouseWheel(sender As Object, e As MouseWheelEventArgs)
        Scv.ScrollToVerticalOffset(-e.Delta)
    End Sub
End Class

