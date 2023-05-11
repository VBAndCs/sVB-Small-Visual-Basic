Imports System.Windows.Controls.Primitives

Friend Class RotateThumb
    Inherits Thumb

    Private centerPoint As Point
    Private startVector As Vector
    Private designerCanvas As Canvas
    Private Pnl As DiagramPanel
    Dim Dsn As Designer
    Dim Item As ListBoxItem
    Dim Diagram As FrameworkElement
    Dim InitialAngle As Double

    Public Property MyInitialAngle As Double

    Public Property RotateAngle As Double
        Get
            Return GetValue(RotateAngleProperty)
        End Get

        Set(value As Double)
            If value >= 360 Then
                value -= 360
            ElseIf value < 0 Then
                value += 360
            End If
            SetValue(RotateAngleProperty, value)
        End Set
    End Property

    Public Shared ReadOnly RotateAngleProperty As DependencyProperty =
                           DependencyProperty.Register("RotateAngle",
                           GetType(Double), GetType(RotateThumb), New PropertyMetadata(AddressOf RotateAngleChanged))

    Shared Sub RotateAngleChanged(rt As RotateThumb, e As DependencyPropertyChangedEventArgs)
        rt.CounterRotateAngle = -e.NewValue - rt.MyInitialAngle
    End Sub

    Public Property CounterRotateAngle As Double
        Get
            Return GetValue(CounterRotateAngleProperty)
        End Get

        Set(value As Double)
            SetValue(CounterRotateAngleProperty, value)
        End Set
    End Property

    Public Shared ReadOnly CounterRotateAngleProperty As DependencyProperty =
                           DependencyProperty.Register("CounterRotateAngle",
                           GetType(Double), GetType(RotateThumb),
                           New PropertyMetadata(0.0))


    Public Property AngleVisibilty As Visibility
        Get
            Return GetValue(AngleVisibiltyProperty)
        End Get

        Set(value As Visibility)
            SetValue(AngleVisibiltyProperty, value)
        End Set
    End Property


    Public Shared ReadOnly AngleVisibiltyProperty As DependencyProperty =
                           DependencyProperty.Register("AngleVisibilty",
                           GetType(Visibility), GetType(RotateThumb),
                           New PropertyMetadata(Visibility.Collapsed))

    Public Sub New()
        AddHandler DragDelta, AddressOf RotateThumb_DragDelta
        AddHandler DragStarted, AddressOf RotateThumb_DragStarted
        RotateAngle = 0
        IsTabStop = False
    End Sub

    Private Sub RotateThumb_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Pnl = Helper.GetDiagramPanel(Me)
        Item = Helper.GetListBoxItem(Pnl)
        Dsn = Helper.GetDesigner(Item)
        designerCanvas = Helper.GetCanvas(Item)
        Diagram = Helper.GetDiagram(Pnl)

    End Sub

    Dim OldState As PropertyState

    Private Sub RotateThumb_DragStarted(sender As Object, e As DragStartedEventArgs)
        OldState = New PropertyState(Pnl.AfterRestoreSub, Diagram, Designer.RotationAngleProperty)

        AngleVisibilty = Windows.Visibility.Visible
        Me.centerPoint = Item.TranslatePoint(New Point(Item.ActualWidth * Item.RenderTransformOrigin.X, Item.ActualHeight * Item.RenderTransformOrigin.Y), Me.designerCanvas)

        Dim startPoint As Point = Mouse.GetPosition(Me.designerCanvas)
        startVector = Point.Subtract(startPoint, Me.centerPoint)

        InitialAngle = Designer.GetRotationAngle(Diagram)
        RotateAngle = InitialAngle
    End Sub

    Private Sub RotateThumb_DragDelta(sender As Object, e As DragDeltaEventArgs)
        Dim currentPoint = Mouse.GetPosition(Me.designerCanvas)
        Dim deltaVector = Point.Subtract(currentPoint, Me.centerPoint)

        Dim angle = Vector.AngleBetween(startVector, deltaVector)
        RotateAngle = Math.Round(InitialAngle + angle, 0)
        Designer.SetRotationAngle(Diagram, RotateAngle)

        Item.InvalidateMeasure()
        Pnl.DiagramGroup?.UpdateSelection()
        Pnl.UpdateLocationBorder()
    End Sub

    Private Sub RotateThumb_DragCompleted(sender As Object, e As DragCompletedEventArgs) Handles Me.DragCompleted
        AngleVisibilty = Visibility.Collapsed
        Dsn.LocationVisibility = Visibility.Collapsed
        ReportChanges()
    End Sub

    Private Sub RotateThumb_PreviewMouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles Me.PreviewMouseDoubleClick
        RotateAngle = 0
        Designer.SetRotationAngle(Diagram, 0)
        Item.InvalidateMeasure()
        Pnl.DiagramGroup?.UpdateSelection()
        Pnl.UpdateLocationBorder()
        e.Handled = True
        ReportChanges()
    End Sub

    Private Sub ReportChanges()
        If OldState.HasChanges Then
            Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
        End If
    End Sub

End Class