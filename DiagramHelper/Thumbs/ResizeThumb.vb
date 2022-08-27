Imports System.Windows.Controls.Primitives

Friend Class ResizeThumb
    Inherits Thumb

    Public Shared Event MeasurementsVisibiltyChanged As EventHandler

    Private Shared _MeasurementsVisibilty As Visibility = Windows.Visibility.Collapsed

    Public Shared Property MeasurementsVisibilty As Visibility
        Get
            Return _MeasurementsVisibilty
        End Get
        Set(value As Visibility)
            _MeasurementsVisibilty = value
            RaiseEvent MeasurementsVisibiltyChanged(Nothing, Nothing)
        End Set
    End Property

    Public Property ResizeAngle As Double
    Private rotateTransform As RotateTransform
    Private skewTransform As SkewTransform
    Private RotateAngle As Double
    Private DeltaX, DeltaY As Double
    Private transformOrigin As Point
    Private Dsn As Designer
    Dim Pnl As DiagramPanel
    Dim Diagram As FrameworkElement
    Dim ResizeCursor As Arrow

    Public Sub New()
        AddHandler DragStarted, AddressOf ResizeThumb_DragStarted
        AddHandler DragDelta, AddressOf ResizeThumb_DragDelta
        AddHandler DragCompleted, AddressOf ResizeThumb_DragCompleted
        Me.Opacity = 0.5
        Me.IsTabStop = False
        Me.RenderTransformOrigin = New Point(0.5, 0.5)
        Me.Cursor = Cursors.None
        Me.ResizeCursor = New Arrow
    End Sub

    Public Property DesignerItem As ListBoxItem
        Get
            Return GetValue(DesignerItemProperty)
        End Get

        Set(value As ListBoxItem)
            SetValue(DesignerItemProperty, value)
        End Set
    End Property

    Public Shared ReadOnly DesignerItemProperty As DependencyProperty =
                           DependencyProperty.Register("DesignerItem",
                           GetType(ListBoxItem), GetType(ResizeThumb))

    Dim OldState As PropertyState

    Private Sub ResizeThumb_DragStarted(sender As Object, e As DragStartedEventArgs)

        ResizeThumb.MeasurementsVisibilty = Windows.Visibility.Visible

        Me.transformOrigin = Pnl.RenderTransformOrigin

        Me.rotateTransform = TryCast(DesignerItem.RenderTransform, RotateTransform)
        If Me.rotateTransform IsNot Nothing Then
            Me.RotateAngle = Me.rotateTransform.Angle * Math.PI / 180.0
        Else
            Me.RotateAngle = 0.0R
        End If

        If Not Double.IsNaN(Diagram.Width) Then Diagram.Width = Double.NaN
        If Not Double.IsNaN(Diagram.Height) Then Diagram.Height = Double.NaN

        Me.skewTransform = TryCast(Diagram.LayoutTransform, SkewTransform)
        If Me.skewTransform Is Nothing Then
            Me.skewTransform = New SkewTransform(0.000000000001, 0.000000000001)
            Me.DeltaX = 0
            Me.DeltaY = 0
            Diagram.LayoutTransform = Me.skewTransform
        Else
            Dim W = Diagram.ActualWidth
            Dim H = Diagram.ActualHeight
            Dim TanAx = Math.Tan(skewTransform.AngleX * Math.PI / 180)
            Dim TanAy = Math.Tan(skewTransform.AngleY * Math.PI / 180)
            Me.DeltaX = (W * TanAy + H) * TanAx / (1 - TanAx * TanAy)
            Me.DeltaY = (W + DeltaX) * TanAy
        End If

        OldState = New PropertyState(Pnl.AfterRestoreSub, Diagram,
                          Designer.FrameWidthProperty, Designer.FrameHeightProperty,
                          Designer.LeftProperty, Designer.TopProperty,
                          FrameworkElement.LayoutTransformProperty)

    End Sub

    Private Sub ResizeThumb_DragDelta(sender As Object, e As DragDeltaEventArgs)
        ResizeThumb.MeasurementsVisibilty = Visibility.Visible
        Dim deltaVertical, deltaHorizontal As Double

        Select Case HorizontalAlignment
            Case HorizontalAlignment.Left
                deltaHorizontal = Math.Min(e.HorizontalChange, Pnl.ActualWidth - Pnl.MinWidth)
                If Math.Abs(deltaHorizontal) * Helper.PxToCm >= 0.1 Then
                    deltaHorizontal = Helper.FixToMm(deltaHorizontal)

                    Canvas.SetTop(DesignerItem, Canvas.GetTop(DesignerItem) + (1 - Me.transformOrigin.X) * deltaHorizontal * Math.Sin(Me.RotateAngle))
                    Canvas.SetLeft(DesignerItem, Canvas.GetLeft(DesignerItem) + deltaHorizontal * Math.Cos(Me.RotateAngle) + (Me.transformOrigin.X * deltaHorizontal * (1 - Math.Cos(Me.RotateAngle))))
                    Pnl.Width = Pnl.ActualWidth - deltaHorizontal

                    If Keyboard.Modifiers = ModifierKeys.Control Then
                        DeltaX -= deltaHorizontal
                        Me.skewTransform.AngleX = Math.Atan2(DeltaX, Pnl.ActualHeight - DeltaY) * 180 / Math.PI
                    End If
                End If

            Case HorizontalAlignment.Right
                deltaHorizontal = Math.Min(-e.HorizontalChange, Pnl.ActualWidth - Pnl.MinWidth)

                If Math.Abs(deltaHorizontal) * Helper.PxToCm >= 0.1 Then
                    deltaHorizontal = Helper.FixToMm(deltaHorizontal)

                    Canvas.SetTop(DesignerItem, Canvas.GetTop(DesignerItem) - Me.transformOrigin.X * deltaHorizontal * Math.Sin(Me.RotateAngle))
                    Canvas.SetLeft(DesignerItem, Canvas.GetLeft(DesignerItem) + (deltaHorizontal * Me.transformOrigin.X * (1 - Math.Cos(Me.RotateAngle))))
                    Pnl.Width = Pnl.ActualWidth - deltaHorizontal

                    If Keyboard.Modifiers = ModifierKeys.Control Then
                        DeltaX += deltaHorizontal
                        Me.skewTransform.AngleX = Math.Atan2(DeltaX, Pnl.ActualHeight - DeltaY) * 180 / Math.PI
                    End If

                End If
        End Select


        Select Case VerticalAlignment
            Case VerticalAlignment.Top
                deltaVertical = Math.Min(e.VerticalChange, Pnl.ActualHeight - Pnl.MinHeight)
                If Math.Abs(deltaVertical) * Helper.PxToCm >= 0.1 Then
                    deltaVertical = Helper.FixToMm(deltaVertical)
                    Canvas.SetTop(DesignerItem, Canvas.GetTop(DesignerItem) + deltaVertical * Math.Cos(-Me.RotateAngle) + (Me.transformOrigin.Y * deltaVertical * (1 - Math.Cos(-Me.RotateAngle))))
                    Canvas.SetLeft(DesignerItem, Canvas.GetLeft(DesignerItem) + deltaVertical * Math.Sin(-Me.RotateAngle) - (Me.transformOrigin.Y * deltaVertical * Math.Sin(-Me.RotateAngle)))
                    Pnl.Height = Pnl.ActualHeight - deltaVertical

                    If Keyboard.Modifiers = ModifierKeys.Control Then
                        DeltaY += deltaVertical
                        Me.skewTransform.AngleY = Math.Atan2(DeltaY, Pnl.ActualWidth + DeltaX) * 180 / Math.PI
                    End If
                End If

            Case VerticalAlignment.Bottom
                deltaVertical = Math.Min(-e.VerticalChange, Pnl.ActualHeight - Pnl.MinHeight)
                If Math.Abs(deltaVertical) * Helper.PxToCm >= 0.1 Then
                    deltaVertical = Helper.FixToMm(deltaVertical)
                    Canvas.SetTop(DesignerItem, Canvas.GetTop(DesignerItem) + (Me.transformOrigin.Y * deltaVertical * (1 - Math.Cos(-Me.RotateAngle))))
                    Canvas.SetLeft(DesignerItem, Canvas.GetLeft(DesignerItem) - deltaVertical * Me.transformOrigin.Y * Math.Sin(-Me.RotateAngle))
                    Pnl.Height = Pnl.ActualHeight - deltaVertical

                    If Keyboard.Modifiers = ModifierKeys.Control Then
                        DeltaY -= deltaVertical
                        Me.skewTransform.AngleY = Math.Atan2(DeltaY, Pnl.ActualWidth + DeltaX) * 180 / Math.PI
                    End If
                End If
        End Select

        Pnl.DiagramGroup?.UpdateSelection()
        Pnl.UpdateLocationBorder()
        e.Handled = True
    End Sub

    Private Sub ResizeThumb_DragCompleted(sender As Object, e As DragCompletedEventArgs)
        ResizeThumb.MeasurementsVisibilty = Visibility.Collapsed
        ReportChanges()
    End Sub

    Private Sub ResizeThumb_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Pnl = Helper.GetDiagramPanel(Me)
        DesignerItem = Helper.GetListBoxItem(Pnl)
        Dsn = Helper.GetDesigner(DesignerItem)
        Diagram = Helper.GetDiagram(Pnl)
    End Sub

    Private Sub ResizeThumb_MouseEnter(sender As Object, e As MouseEventArgs) Handles Me.MouseEnter
        Me.rotateTransform = TryCast(DesignerItem.RenderTransform, RotateTransform)
        Dim TotalAngle = Me.ResizeAngle + If(Me.rotateTransform Is Nothing, 0, Me.rotateTransform.Angle)

        If Me.Cursor Is Cursors.None OrElse Me.ResizeCursor.Angle <> TotalAngle Then
            Me.ResizeCursor.Angle = TotalAngle
            Me.Cursor?.Dispose()
            Me.Cursor = CursorHelper.CreateCursor(Me.ResizeCursor, -1, -1)
        End If
    End Sub

    Private Sub ResizeThumb_PreviewMouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles Me.PreviewMouseDoubleClick

        If VerticalAlignment = VerticalAlignment.Bottom OrElse VerticalAlignment = VerticalAlignment.Top Then
            Pnl.Width = Pnl.Height
        ElseIf HorizontalAlignment = Windows.HorizontalAlignment.Left OrElse HorizontalAlignment = Windows.HorizontalAlignment.Right Then
            Pnl.Height = Pnl.Width
        End If
        Helper.UpdateControl(Pnl)
        Pnl.DiagramGroup?.UpdateSelection()
        Pnl.UpdateLocationBorder()
        ReportChanges()
    End Sub

    Sub ReportChanges()
        If OldState?.HasChanges Then
            Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Try
            Me.Cursor?.Dispose()
        Catch
        End Try

        Try
            MyBase.Finalize()
        Catch
        End Try
    End Sub
End Class
