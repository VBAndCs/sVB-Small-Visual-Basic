Public Class ToolBox

    Dim tabsLoaded As Boolean

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

    Sub AddTab(FolderName As String)
        Dim Expan As New Expander With {.IsExpanded = IsFirstTab}
        Expan.Style = Me.FindResource("ExpanderStyle")
        IsFirstTab = False
        Expan.Header = IO.Path.GetFileNameWithoutExtension(FolderName)

        Dim WrpPnl As New WrapPanel With {.Background = Brushes.White}
        Dim Scv As New ScrollViewer With {.VerticalScrollBarVisibility = ScrollBarVisibility.Auto, .MaxHeight = 120, .Margin = New Thickness(1), .Content = WrpPnl}
        Expan.Content = Scv
        Dim Files = IO.Directory.GetFiles(FolderName, "*.xaml")
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
                    Catch ex As Exception

                    End Try
                Loop While WrpPnl2.Children.Count > 0

            Else
                Try
                    Dim diagram = Helper.CreateDiagram(FileName)
                    Dim Item As New ToolBoxItem(diagram, FileName)
                    WrpPnl.Children.Add(Item)
                Catch ex As Exception

                End Try

            End If
        Next
        ToolBoxTabs.Children.Add(Expan)
    End Sub
End Class

