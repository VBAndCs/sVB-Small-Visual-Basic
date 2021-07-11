Imports System.Windows.Controls.Primitives
Imports System.Globalization

Public Class DiagramObject
    Dim WithEvents Diagram As FrameworkElement
    Dim Pnl As DiagramPanel
    Dim DesignerItem As ListBoxItem
    Dim Dsn As Designer
    Dim Canv As Canvas
    Dim Scv As ScrollViewer

    Friend Shared Diagrams As New Dictionary(Of FrameworkElement, DiagramObject)

    Dim TmpIsSelected As Boolean
    Dim DraggingDiagram As Boolean = False
    Dim DraggingStartPoint As Point
    Dim OldPropertyState As PropertyState
    Dim Unit As UndoRedoUnit


    Private Sub New(Diagram As FrameworkElement)
        Me.Diagram = Diagram
        Canv = Helper.GetCanvas(Diagram)
        Scv = Helper.GetScrollViewer(Canv)
        Dsn = Helper.GetDesigner(Scv)
    End Sub

    Shared Function CreateDiagramObject(Diagram As FrameworkElement) As DiagramObject
        Dim DiagramObj As DiagramObject

        If Diagrams.ContainsKey(Diagram) Then
            DiagramObj = Diagrams(Diagram)
        Else
            DiagramObj = New DiagramObject(Diagram)
            Diagrams.Add(Diagram, DiagramObj)
        End If

        DiagramObj.Pnl = Helper.GetDiagramPanel(Diagram)
        DiagramObj.DesignerItem = Helper.GetListBoxItem(DiagramObj.Pnl)

        Return DiagramObj
    End Function

    Friend Sub AfterRestoreAction()
        Helper.UpdateControl(Diagram)
        Pnl.AdjustConnectors()

        If Keyboard.FocusedElement IsNot DesignerItem Then
            If Not DesignerItem.IsSelected Then
                Dsn.SelectedIndex = -1
                DesignerItem.IsSelected = True
            End If
            DesignerItem.Focus()
        End If

        If Pnl.DiagramGroup IsNot Nothing Then Pnl.DiagramGroup.Select()

        Dsn.ScrollIntoView(Diagram)
        Helper.UpdateControl(Pnl)
        Pnl.OnConnectorsPositionChangd()
    End Sub

    Private Sub UpdateLocationBorder()
        Pnl.ConnectorsGrid.LayoutTransform = Diagram.LayoutTransform
        Pnl.FocusRectangle.LayoutTransform = Diagram.LayoutTransform
        Pnl.FocusRectangle.Width = Diagram.ActualWidth
        Pnl.FocusRectangle.Height = Diagram.ActualHeight

        Pnl.FocusRectangle.StrokeThickness = 2
        Dsn.LocationVisibility = Windows.Visibility.Visible
        UpdateTbLocation()
    End Sub

    Sub UpdateTbLocation()
        Dim P As Point = GetLeftTopPoint()
        Dsn.TbLeftLocation.Text = P.X.ToString("F1")
        Dsn.TbTopLocation.Text = P.Y.ToString("F1")
        Helper.UpdateControl(Dsn.TbLeftLocation)
        Helper.UpdateControl(Dsn.TbTopLocation)

        'Dim Location = GetLeftTopPoint(Dsn, False)
        Dim x = P.X * Helper.CmToPx * Dsn.Scale - 5
        Dim y = P.Y * Helper.CmToPx * Dsn.Scale - 5
        Dsn.TbLeftLocation.Margin = New Thickness(x - Dsn.TbLeftLocation.ActualWidth * Dsn.Scale, y + 2, 0, 0)
        Dsn.TbTopLocation.Margin = New Thickness(x + 2, y - Dsn.TbTopLocation.ActualHeight * Dsn.Scale, 0, 0)

    End Sub

    Friend Function GetLeftTopPoint(Optional RelativeTo As FrameworkElement = Nothing, Optional InCm As Boolean = True) As Point
        Dim P As Point
        Dim Lt = Pnl.FocusRectangle.LayoutTransform
        If Lt IsNot Nothing Then P = Lt.Transform(New Point(0, 0))
        Dim Rt = DesignerItem.RenderTransform
        If Rt IsNot Nothing Then P = Rt.Transform(P)
        If RelativeTo Is Nothing Then
            P = Pnl.FocusRectangle.TransformToVisual(Canv).Transform(P)
        Else
            P = Pnl.FocusRectangle.TransformToVisual(RelativeTo).Transform(P)
        End If

        If InCm Then
            P.X = Math.Round(P.X * Helper.PxToCm, 2)
            P.Y = Math.Round(P.Y * Helper.PxToCm, 2)
        End If
        Return P
    End Function

    Private Sub Diagram_MouseLeave(sender As Object, e As MouseEventArgs) Handles Diagram.MouseLeave
        If DraggingDiagram Then Diagram_PreviewMouseLeftButtonUp(Nothing, Nothing)
    End Sub

    Private Sub Diagram_GotFocus(ByVal sender As Object, ByVal e As EventArgs) Handles Diagram.GotFocus
        If Keyboard.Modifiers = ModifierKeys.Control Then
            Pnl.IsSelected = Not Pnl.IsSelected
        ElseIf Pnl.IsSelected = False Then
            Connection.DeselectAll(Dsn)
            Dsn.SelectedIndex = -1
            Pnl.IsSelected = True
        End If
    End Sub

    Friend Sub MoveDiagram(HorizontalChange As Double, VerticalChange As Double, Optional IgnoreRotation As Boolean = False)

        Dim dragDelta As New Point(HorizontalChange, VerticalChange)
        If Not IgnoreRotation Then
            Dim rotateTransform As RotateTransform = TryCast(DesignerItem.RenderTransform, RotateTransform)
            If rotateTransform IsNot Nothing Then dragDelta = rotateTransform.Transform(dragDelta)
        End If

        Dim OldLeft = Dsn.SelectedBounds.Left  'Canvas.GetLeft(DesignerItem)
        If Double.IsNaN(OldLeft) Then OldLeft = 0
        Dim OldTop = Dsn.SelectedBounds.Top  'Canvas.GetTop(DesignerItem)
        If Double.IsNaN(OldTop) Then OldTop = 0

        Dim NewLeft As Double = OldLeft
        If Math.Abs(dragDelta.X) * Helper.PxToCm >= 0.1 Then
            'Dim MarginX = (LstDesigner.SelectedBounds.Width - Pnl.ActualWidth) / 2
            NewLeft = Math.Max(OldLeft + dragDelta.X, 1)
            NewLeft = Math.Min(Canv.ActualWidth - Dsn.SelectedBounds.Width - 1, NewLeft)
            NewLeft = Helper.FixToMm(NewLeft)
        End If

        Dim NewTop As Double = OldTop
        If Math.Abs(dragDelta.Y) * Helper.PxToCm >= 0.1 Then
            'Dim MarginY = (LstDesigner.SelectedBounds.Height - Pnl.ActualHeight) / 2
            NewTop = Math.Max(OldTop + dragDelta.Y, 1)
            NewTop = Math.Min(Canv.ActualHeight - Dsn.SelectedBounds.Height - 1, NewTop)
            NewTop = Helper.FixToMm(NewTop)
        End If

        ' Agjust TopLeft Point to mm
        Dim P = GetLeftTopPoint()
        If HorizontalChange > 0 Then
            P.X = Math.Ceiling(P.X * 10) / 10 - P.X
        Else
            P.X = Math.Floor(P.X * 10) / 10 - P.X
        End If

        NewLeft += P.X * Helper.CmToPx

        If VerticalChange > 0 Then
            P.Y = Math.Ceiling(P.Y * 10) / 10 - P.Y
        Else
            P.Y = Math.Floor(P.Y * 10) / 10 - P.Y
        End If
        NewTop += P.Y * Helper.CmToPx

        Dim DeltaX = NewLeft - OldLeft
        Dim DeltaY = NewTop - OldTop

        If Not DesignerItem.IsSelected Then
            Canvas.SetLeft(DesignerItem, Canvas.GetLeft(DesignerItem) + DeltaX)
            Canvas.SetTop(DesignerItem, Canvas.GetTop(DesignerItem) + DeltaY)
            Pnl.OnConnectorsPositionChangd()
        End If

        For Each D In Dsn.SelectedItems
            Dim Item = Helper.GetListBoxItem(D)
            Canvas.SetLeft(Item, Canvas.GetLeft(Item) + DeltaX)
            Canvas.SetTop(Item, Canvas.GetTop(Item) + DeltaY)
            Helper.GetDiagramPanel(D).OnConnectorsPositionChangd()
        Next

        Dim scale = TryCast(Canv.LayoutTransform, ScaleTransform)
        Dim F = If(scale IsNot Nothing, scale.ScaleX, 1)
        Dim Delay = 0

        If NewLeft * F <= Scv.HorizontalOffset OrElse
              ((NewLeft + Dsn.SelectedBounds.Width) * F) >= (Scv.HorizontalOffset + Scv.ViewportWidth) Then
            Delay = 40
            Scv.ScrollToHorizontalOffset(Scv.HorizontalOffset + DeltaX * F)
        End If

        If NewTop * F <= Scv.VerticalOffset OrElse
                  ((NewTop + Dsn.SelectedBounds.Height) * F) >= (Scv.VerticalOffset + Scv.ViewportHeight) Then
            Delay = 40
            Scv.ScrollToVerticalOffset(Scv.VerticalOffset + DeltaY * F)
        End If
        Dsn.SelectedBounds.X = NewLeft
        Dsn.SelectedBounds.Y = NewTop

        UpdateTbLocation()
        Helper.UpdateControl(Dsn)

        If Delay > 0 Then Helper.Delay(Delay)
    End Sub

    Private Sub Diagram_PreviewMouseMove(sender As Object, e As MouseEventArgs) Handles Diagram.PreviewMouseMove
        If e.LeftButton = MouseButtonState.Released Then Return

        If DraggingDiagram Then
            Dim p = e.GetPosition(Pnl)
            Dim x = p.X - DraggingStartPoint.X
            Dim y = p.Y - DraggingStartPoint.Y
            MoveDiagram(x, y)

            e.Handled = True
        End If
    End Sub

    Private Sub Diagram_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles Diagram.PreviewMouseDown
        If Dsn.ConnectionMode = False Then Dsn.ConnectionMode = (Keyboard.Modifiers = ModifierKeys.Shift)

        If Dsn.ConnectionMode Then
            If Dsn.ConnectionSourceDiagram Is Nothing Then
                Dsn.ConnectionSourceDiagram = Diagram
                Dsn.SourceConnector = New ConnectorThumb(Diagram)
                Dsn.SourceConnector.StartDrawConnection()
            ElseIf Dsn.ConnectionSourceDiagram Is Diagram Then
                Return
            Else
                Dsn.ConnectionTargetDiagram = Diagram
                Dsn.TargetConnector = New ConnectorThumb(Diagram)
                Dsn.SourceConnector.EndDrawConnection()
            End If
        End If

        Diagram.CaptureMouse()
        DraggingDiagram = True

        If DesignerItem.IsSelected Then
            Dsn.SelectedBounds = Dsn.GetSelectionBounds()
            Diagram_GotFocus(Nothing, Nothing)
        Else
            Diagram_GotFocus(Nothing, Nothing)
            Dsn.SelectedBounds = Dsn.GetSelectionBounds()
        End If

        UpdateLocationBorder()

        DraggingStartPoint = e.GetPosition(Pnl)

        Pnl.ExitIsSelectedChanged = True
        TmpIsSelected = Pnl.IsSelected
        DesignerItem.Focus()
        Pnl.IsSelected = False
        OldPropertyState = New PropertyState(AddressOf AfterRestoreAction, Diagram, Designer.LeftProperty, Designer.TopProperty)
        Unit = New UndoRedoUnit(OldPropertyState)

        For Each D As FrameworkElement In Dsn.SelectedItems
            If D Is Diagram Then Continue For
            Dim Ps As New PropertyState(AddressOf AfterRestoreAction, D, OldPropertyState.Keys.ToArray)
            Unit.Add(Ps)
        Next
        e.Handled = True
    End Sub

    Private Sub Diagram_PreviewMouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Diagram.PreviewMouseLeftButtonUp
        Diagram.ReleaseMouseCapture()
        DraggingDiagram = False
        DesignerItem.Focus()
        Pnl.IsSelected = TmpIsSelected
        Pnl.ExitIsSelectedChanged = False
        Dsn.LocationVisibility = Windows.Visibility.Collapsed

        If OldPropertyState IsNot Nothing AndAlso OldPropertyState.HasChanges Then
            For Each State As PropertyState In Unit
                State.SetNewValue()
            Next
            Dsn.UndoStack.ReportChanges(Unit)
        End If
    End Sub


#Region "Editor"

    Friend Sub BeginEdit()
        Designer.Editing = True
        Dim P = GetLeftTopPoint(InCm:=False)
        Dsn.Editor.Width = Diagram.ActualWidth
        Dsn.Editor.Height = Diagram.ActualHeight
        Dsn.Editor.Text = Pnl.DiagramTextBlock.Text

        Dim Pup As Popup = Dsn.Editor.Parent
        Pup.PlacementTarget = Diagram
        Pup.Placement = PlacementMode.Relative
        Pup.IsOpen = True
        Dsn.Editor.Focus()
        Dsn.Editor.SelectAll()

        AddHandler Dsn.Editor.PreviewKeyDown, AddressOf Editor_PreviewKeyDown
        AddHandler Pup.Closed, AddressOf Pup_Closed
    End Sub

    Friend Sub EndEdit(Commit As Boolean)
        Designer.Editing = False
        RemoveHandler Dsn.Editor.PreviewKeyDown, AddressOf Editor_PreviewKeyDown
        If Commit Then
            Dim OldState As New PropertyState(Pnl.DiagramTextBlock, TextBlock.TextProperty)
            Pnl.DiagramTextBlock.Text = Dsn.Editor.Text
            Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
        End If

        Dim Pup As Popup = Dsn.Editor.Parent
        RemoveHandler Dsn.Editor.PreviewKeyDown, AddressOf Editor_PreviewKeyDown
        RemoveHandler Pup.Closed, AddressOf Pup_Closed
        Pup.IsOpen = False
        DesignerItem.Focus()
    End Sub

    Private Sub Editor_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            If Keyboard.Modifiers = ModifierKeys.Control Then
                Dsn.Editor.SelectedText = vbCrLf
                Dsn.Editor.SelectionLength = 0
                Dsn.Editor.SelectionStart += 2
            Else
                EndEdit(True)
            End If
        ElseIf e.Key = Key.Escape Then
            EndEdit(False)
        End If
    End Sub

    Private Sub Pup_Closed(sender As Object, e As EventArgs)
        EndEdit(True)
    End Sub

    Friend Sub DiagramTextBlock_TextChanged()
        Designer.SetDiagramText(Diagram, Pnl.DiagramTextBlock.Text)
    End Sub

#End Region


    Sub OutlineText()
        Dim Tb = Pnl.DiagramTextBlock
        Dim TextForeground = Designer.GetDiagramTextforeground(Diagram)        
        Dim TextBackground = Designer.GetDiagramTextBackground(Diagram)

        Dim Gd1 As New GeometryDrawing(TextBackground, Nothing,
                                   New RectangleGeometry(New Rect(0, 0, Tb.ActualWidth, Tb.ActualHeight)))

        Dim Tf = New Typeface(Tb.FontFamily, Tb.FontStyle, Tb.FontWeight, Tb.FontStretch)
        Dim F As New FormattedText(Tb.Text, CultureInfo.CurrentCulture, Tb.FlowDirection,
                                   Tf, Tb.FontSize, TextForeground)
        Dim Gd2 As New GeometryDrawing(Designer.GetDiagramTextOutlineFill(Diagram), New Pen(TextForeground, Designer.GetDiagramTextOutlineThickness(Diagram)), F.BuildGeometry(New Point))
        Dim Dg As New DrawingGroup
        Dg.Children.Add(Gd1)
        Dg.Children.Add(Gd2)
        Dim TextBursh As New DrawingBrush(Dg)
        Tb.Background = TextBursh
        Tb.Foreground = Nothing
    End Sub

    Private Sub Diagram_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Diagram.PreviewMouseLeftButtonDown
        If e.ClickCount > 1 Then
            Dsn.OnDiagramDoubleClick(Diagram)
        End If
    End Sub
End Class
