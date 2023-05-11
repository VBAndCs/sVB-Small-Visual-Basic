Friend Class ToolBoxItem
    Inherits Border

    Sub New(Diagram As UIElement, Optional FileName As String = "")
        MyBase.New()

        Me.Name = IO.Path.GetFileNameWithoutExtension(FileName)
        Dim Fw = TryCast(Diagram, FrameworkElement)

        If Fw IsNot Nothing Then
            Me.Width = Math.Max(34, Fw.MinWidth + Fw.Margin.Left + Fw.Margin.Right + 6)
            Me.Height = Math.Max(34, Fw.MinHeight + Fw.Margin.Top + Fw.Margin.Bottom + 6)
            If Fw.ToolTip = "" Then
                Fw.ToolTip = Me.Name
            Else
                Me.Name = Fw.ToolTip.ToString().Replace(" ", "_")
            End If
        Else
            Me.Width = 34
            Me.Height = 34
            If Me.Name = "" Then Me.Name = "Diagram"
        End If

        Me.Padding = New Thickness(2)
        Me.Child = Diagram
        Me.Margin = New Thickness(3)
        Me.BorderBrush = Brushes.Black
    End Sub


    Dim _isSelected As Boolean
    Dim except As ToolBoxItem

    Public Property IsSelected As Boolean
        Get
            Return _isSelected
        End Get

        Set(value As Boolean)
            If except Is Me Then
                except = Nothing
                Return
            End If

            If value Then
                except = Me
                ToolBox.DeSelectAll()
                _isSelected = True
                Me.BorderThickness = New Thickness(2)
                ToolBox.SelectedItem = Me
            Else
                _isSelected = False
                ToolBox.SelectedItem = Nothing
                Me.BorderThickness = New Thickness(
                    If(Me.BorderBrush Is Brushes.Brown, 1, 0)
                )
            End If
        End Set
    End Property

    Private Sub ToolBoxItem_MouseEnter(sender As Object, e As MouseEventArgs) Handles Me.MouseEnter
        If Not _isSelected Then Me.BorderThickness = New Thickness(1)
        Me.BorderBrush = Brushes.Brown
    End Sub

    Private Sub ToolBoxItem_MouseLeave(sender As Object, e As MouseEventArgs) Handles Me.MouseLeave
        If Not (IsDragging OrElse _isSelected) Then Me.BorderThickness = New Thickness(0)
        Me.BorderBrush = Brushes.Black
    End Sub

    Private Sub ToolBoxItem_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseLeftButtonDown
        Me.IsSelected = Not _isSelected
    End Sub

    Dim IsDragging As Boolean = False
    Friend ToolBox As ToolBox

    Private Sub ToolBoxItem_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If e.LeftButton = MouseButtonState.Pressed Then
            IsDragging = True
            Dim x = DragDrop.DoDragDrop(Me, Me, DragDropEffects.Copy)
            IsDragging = False
            Dim pos = e.GetPosition(Me)
            If _isSelected AndAlso (pos.X < 0 OrElse pos.Y < 0 OrElse pos.X > Me.ActualWidth OrElse pos.Y > Me.ActualWidth) Then
                Me.IsSelected = False
            End If

            If Not _isSelected Then Me.BorderThickness = New Thickness(0)
        End If
    End Sub
End Class
