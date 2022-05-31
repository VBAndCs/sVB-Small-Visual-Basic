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
    Private Dsn As Designer


    Public Sub New()
        AddHandler DragStarted, AddressOf ResizeThumb_DragStarted
        AddHandler DragDelta, AddressOf ResizeThumb_DragDelta
        AddHandler DragCompleted, AddressOf ResizeThumb_DragCompleted
        Me.Opacity = 0.5
        Me.IsTabStop = False
    End Sub

    Sub Thimb_loaded() Handles Me.Loaded
        Me.Cursor = CursorHelper.CreateCursor(New Arrow(Me.ResizeAngle), -1, -1)
    End Sub

    Dim OldState As PropertyState

    Private Sub ResizeThumb_DragStarted(ByVal sender As Object, ByVal e As DragStartedEventArgs)

        DsnResizeThumb.MeasurementsVisibilty = Windows.Visibility.Visible

        OldState = New PropertyState(Dsn,
                          Designer.PageWidthProperty,
                          Designer.PageHeightProperty)

    End Sub

    Private Sub ResizeThumb_DragDelta(ByVal sender As Object, ByVal e As DragDeltaEventArgs)
        DsnResizeThumb.MeasurementsVisibilty = Visibility.Visible
        Dim deltaVertical, deltaHorizontal As Double

        If HorizontalAlignment = HorizontalAlignment.Right Then
            deltaHorizontal = Math.Min(-e.HorizontalChange, Dsn.DesignerCanvas.ActualWidth - 150)

            If Math.Abs(deltaHorizontal) * Helper.PxToCm >= 0.1 Then
                deltaHorizontal = Helper.FixToMm(deltaHorizontal)
                Dsn.PageWidth = Dsn.DesignerCanvas.ActualWidth - deltaHorizontal
            End If
        End If


        Select Case VerticalAlignment
            Case VerticalAlignment.Bottom
                deltaVertical = Math.Min(-e.VerticalChange, Dsn.DesignerCanvas.ActualHeight - 150)
                If Math.Abs(deltaVertical) * Helper.PxToCm >= 0.1 Then
                    deltaVertical = Helper.FixToMm(deltaVertical)
                    Dsn.PageHeight = Dsn.DesignerCanvas.ActualHeight - deltaVertical
                End If
        End Select

        e.Handled = True
    End Sub

    Private Sub ResizeThumb_DragCompleted(ByVal sender As Object, ByVal e As DragCompletedEventArgs)
        DsnResizeThumb.MeasurementsVisibilty = Visibility.Collapsed
        ReportChanges()
    End Sub

    Private Sub ResizeThumb_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dsn = Helper.GetDesigner(sender)
    End Sub

    Sub ReportChanges()
        If OldState.HasChanges Then
            Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
        End If
    End Sub

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
