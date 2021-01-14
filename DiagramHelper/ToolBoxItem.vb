Friend Class ToolBoxItem
    Inherits Border

    Sub New(Diagram As UIElement, Optional FileName As String = "")
        MyBase.New()

        Dim Fw = TryCast(Diagram, FrameworkElement)
        If Fw IsNot Nothing Then
            Me.Width = Math.Max(34, Fw.MinWidth + Fw.Margin.Left + Fw.Margin.Right + 6)
            Me.Height = Math.Max(34, Fw.MinHeight + Fw.Margin.Top + Fw.Margin.Bottom + 6)
            If Fw.ToolTip = "" Then Fw.ToolTip = IO.Path.GetFileNameWithoutExtension(FileName)
        Else
            Me.Width = 34
            Me.Height = 34
        End If

        Me.Padding = New Thickness(2)
        Me.Child = Diagram
        Me.Margin = New Thickness(3)
        Me.BorderBrush = Brushes.Black
    End Sub

    Private Sub ToolBoxItem_MouseEnter(sender As Object, e As MouseEventArgs) Handles Me.MouseEnter
        Me.BorderThickness = New Thickness(1)
    End Sub

    Private Sub ToolBoxItem_MouseLeave(sender As Object, e As MouseEventArgs) Handles Me.MouseLeave
        If Not IsDragging Then Me.BorderThickness = New Thickness(0)
    End Sub

    Private Sub ToolBoxItem_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseLeftButtonDown
        Me.BorderThickness = New Thickness(2)
    End Sub

    Private Sub ToolBoxItem_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseLeftButtonUp
        Me.BorderThickness = New Thickness(1)
    End Sub

    Dim IsDragging As Boolean = False

    Private Sub ToolBoxItem_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If e.LeftButton = MouseButtonState.Pressed Then
            IsDragging = True
            Dim x = DragDrop.DoDragDrop(Me, Me, DragDropEffects.Copy)
            IsDragging = False
            Me.BorderThickness = New Thickness(0)
        End If
    End Sub
End Class
