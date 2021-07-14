Imports System.Windows.Controls.Primitives

Friend Class DsnResizeThumb
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
    Private DeltaX, DeltaY As Double
    Private adorner As Adorner
    Private Dsn As Designer
    Dim ResizeCursor As Arrow

    Public Sub New()
        AddHandler DragStarted, AddressOf ResizeThumb_DragStarted
        AddHandler DragDelta, AddressOf ResizeThumb_DragDelta
        AddHandler DragCompleted, AddressOf ResizeThumb_DragCompleted
        Me.Opacity = 0.5
        Me.IsTabStop = False
        Me.Cursor = Cursors.None
        Me.ResizeCursor = New Arrow
    End Sub

    Dim OldState As PropertyState

    Private Sub ResizeThumb_DragStarted(ByVal sender As Object, ByVal e As DragStartedEventArgs)

        ResizeThumb.MeasurementsVisibilty = Windows.Visibility.Visible

        OldState = New PropertyState(Designer.FrameWidthProperty,
                          Designer.FrameHeightProperty,
                          Designer.LeftProperty, Designer.TopProperty,
                          FrameworkElement.LayoutTransformProperty)

    End Sub

    Private Sub ResizeThumb_DragDelta(ByVal sender As Object, ByVal e As DragDeltaEventArgs)
        ResizeThumb.MeasurementsVisibilty = Visibility.Visible
        Dim deltaVertical, deltaHorizontal As Double

        Select Case HorizontalAlignment
            Case HorizontalAlignment.Right
                deltaHorizontal = Math.Min(-e.HorizontalChange, Dsn.ActualWidth - Dsn.MinWidth)

                If Math.Abs(deltaHorizontal) * Helper.PxToCm >= 0.1 Then
                    deltaHorizontal = Helper.FixToMm(deltaHorizontal)

                    Dsn.Width = Dsn.ActualWidth - deltaHorizontal

                End If
        End Select


        Select Case VerticalAlignment
            Case VerticalAlignment.Bottom
                deltaVertical = Math.Min(-e.VerticalChange, Dsn.ActualHeight - Dsn.MinHeight)
                If Math.Abs(deltaVertical) * Helper.PxToCm >= 0.1 Then
                    deltaVertical = Helper.FixToMm(deltaVertical)
                    Dsn.Height = Dsn.ActualHeight - deltaVertical

                End If
        End Select

        e.Handled = True
    End Sub

    Private Sub ResizeThumb_DragCompleted(ByVal sender As Object, ByVal e As DragCompletedEventArgs)
        ResizeThumb.MeasurementsVisibilty = Visibility.Collapsed
        ReportChanges()
    End Sub

    Private Sub ResizeThumb_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dsn = Helper.GetDesigner(sender)
    End Sub

    Private Sub ResizeThumb_MouseEnter(sender As Object, e As MouseEventArgs) Handles Me.MouseEnter

        If Me.Cursor Is Cursors.None Then
            Try
                Me.Cursor.Dispose()
            Catch
            End Try
            Me.Cursor = CursorHelper.CreateCursor(Me.ResizeCursor, -1, -1)
        End If
    End Sub

    Private Sub ResizeThumb_PreviewMouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles Me.PreviewMouseDoubleClick

        If VerticalAlignment = VerticalAlignment.Bottom OrElse VerticalAlignment = VerticalAlignment.Top Then
            Dsn.Width = Dsn.Height
        ElseIf HorizontalAlignment = Windows.HorizontalAlignment.Left OrElse HorizontalAlignment = Windows.HorizontalAlignment.Right Then
            Dsn.Height = Dsn.Width
        End If
        Helper.UpdateControl(Dsn)
        ReportChanges()
    End Sub

    Sub ReportChanges()
        If OldState.HasChanges Then
            Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
        End If
    End Sub

    'Private Sub ResizeThumb_MouseMove(sender As Object, e As MouseEventArgs) Handles DesignerCanvas.MouseMove
    '    Dim pos = e.GetPosition(DesignerCanvas)
    '    If pos.X > DesignerCanvas.Width - 3 Then
    '        If DesignerCanvas.Cursor Is Nothing Then
    '            DesignerCanvas.Cursor = CursorHelper.CreateCursor(Me.ResizeCursor, -1, -1)
    '        ElseIf DesignerCanvas.Cursor Is Cursors.None Then
    '            Try
    '                Me.Cursor?.Dispose()
    '            Catch : End Try
    '            DesignerCanvas.Cursor = CursorHelper.CreateCursor(Me.ResizeCursor, -1, -1)
    '        End If

    '    Else : Try
    '            Me.Cursor?.Dispose()
    '            DesignerCanvas.Cursor = Nothing
    '        Catch : End Try
    '    End If
    'End Sub

    Protected Overrides Sub Finalize()
        Try
            Me.Cursor.Dispose()
        Catch

        End Try

        Try
            MyBase.Finalize()
        Catch

        End Try
    End Sub
End Class
