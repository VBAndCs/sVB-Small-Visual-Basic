Public Class DesignerDecorator

    Function GetDesigner(Mi As Control) As Designer
        Dim Cntx As ContextMenu = Mi.Parent
        Return Helper.GetDesigner(Cntx.PlacementTarget)
    End Function

    Friend Sub ContextMenu_Opened(sender As Object, e As RoutedEventArgs)
        Dim Cntx As ContextMenu = sender
        Dim Dsn = Helper.GetDesigner(Cntx.PlacementTarget)
        For Each m In Cntx.Items
            Dim Mi = TryCast(m, MenuItem)
            If Mi IsNot Nothing Then
                Select Case Mi.Header
                    Case "Save"
                        Mi.IsEnabled = Dsn.HasChanges
                    Case "Save As..."
                        Mi.IsEnabled = Dsn.FileName <> ""
                    Case "Save To Image", "Print"
                        Mi.IsEnabled = Dsn.Items.Count > 0
                    Case "Paste"
                        Mi.IsEnabled = Dsn.CanPaste
                    Case "Undo"
                        Mi.IsEnabled = Dsn.CanUndo
                    Case "Redo"
                        Mi.IsEnabled = Dsn.CanRedo
                    Case "Close"
                        Mi.IsEnabled = Not Dsn.IsNew OrElse Dsn.Pages.Count > 1
                End Select
            End If
        Next

    End Sub


    Friend Sub SaveMenuItem_Click(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).Save()
    End Sub

    Friend Sub SaveAsMenuItem_Click(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).SaveAs()
    End Sub

    Private Sub NewMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Designer.OpenNewPage()
    End Sub

    Private Sub CloseMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Designer.ClosePage()
    End Sub

    Private Sub OpenMenuItem_Click(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).Open()
    End Sub

    Private Sub SaveImageMenuItem_Click(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).SaveToImage()
    End Sub

    Private Sub PrintMenuItem_Click(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).Print()
    End Sub

    Private Sub SellectAllMenuItem_Click(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).SelectAll()
    End Sub

    Private Sub PasteMenuItem_Click(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).Paste()
    End Sub

    Private Sub RedoMenuItem_Click(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).Redo()
    End Sub

    Private Sub UndoMenuItem_Click(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).Undo()
    End Sub

    Private Sub ShowGridMenuItem_Checked(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).ShowGrid = True
    End Sub

    Private Sub ShowGridMenuItem_Unchecked(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).ShowGrid = False
    End Sub

    Private Sub GridBrushMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.ChangeBrush(GetDesigner(sender).GridPen, Pen.BrushProperty)
    End Sub

    Private Sub PageBackgroundMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.ChangeBrush(GetDesigner(sender).DesignerCanvas, Border.BackgroundProperty)
    End Sub

    Private Sub DecreaseThicknessMenuItem_Click(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).IncreaseGridThickness(-0.1)
    End Sub

    Private Sub IncreaseThicknessMenuItem_Click(sender As Object, e As RoutedEventArgs)
        GetDesigner(sender).IncreaseGridThickness(0.1)
    End Sub

    Private Sub PageSizeMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Dim Dsn = GetDesigner(sender)
        Dim WndPageSetup As New WndPageSetup
        WndPageSetup.PageWidth = Dsn.PageWidth / Helper.CmToPx
        WndPageSetup.PageHeight = Dsn.PageHeight / Helper.CmToPx
        If WndPageSetup.ShowDialog = True Then
            Dim OldState As New PropertyState(Dsn, Designer.PageWidthProperty, Designer.PageHeightProperty)
            Dsn.PageWidth = WndPageSetup.PageWidth * Helper.CmToPx
            Dsn.PageHeight = WndPageSetup.PageHeight * Helper.CmToPx
            If OldState.HasChanges Then
                Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
            End If
        End If
    End Sub
End Class
