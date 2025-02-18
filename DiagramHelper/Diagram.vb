﻿Imports System.Windows.Controls.Primitives

Public Class DiagramObject
    Dim WithEvents Diagram As FrameworkElement
    Dim Pnl As DiagramPanel
    Dim DesignerItem As ListBoxItem
    Dim Dsn As Designer
    Dim Canv As Canvas
    Dim Scv As ScrollViewer
    Dim DraggingDiagram As Boolean = False
    Dim DraggingStartPoint As Point
    Dim OldPropertyState As PropertyState
    Dim Unit As UndoRedoUnit
    Friend Shared Diagrams As New Dictionary(Of FrameworkElement, DiagramObject)

    Private Sub New(Diagram As FrameworkElement)
        Me.Diagram = Diagram
        Diagram.AddHandler(
              UIElement.PreviewMouseDownEvent,
              New RoutedEventHandler(AddressOf Diagram_PreviewMouseDown),
              True
         )

        Diagram.AddHandler(
              UIElement.PreviewMouseLeftButtonDownEvent,
              New RoutedEventHandler(AddressOf Diagram_PreviewMouseLeftButtonDown),
              True
         )

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
        Pnl.DiagramGroup?.UpdateSelection()
        Pnl.UpdateLocationBorder()
        Dsn.LocationVisibility = Windows.Visibility.Collapsed
    End Sub

    Friend Function GetLeftTopPoint(Optional relativeTo As FrameworkElement = Nothing, Optional inCm As Boolean = True) As Point
        Dim p As Point
        Dim Lt = Pnl.FocusRectangle.LayoutTransform
        If Lt IsNot Nothing Then p = Lt.Transform(New Point(0, 0))

        Dim rt = DesignerItem.RenderTransform
        If rt IsNot Nothing Then p = rt.Transform(p)

        If relativeTo Is Nothing Then
            p = Pnl.FocusRectangle.TransformToVisual(Canv).Transform(p)
        Else
            p = Pnl.FocusRectangle.TransformToVisual(relativeTo).Transform(p)
        End If

        If inCm Then
            p.X = Math.Round(p.X * Helper.PxToCm, 2)
            p.Y = Math.Round(p.Y * Helper.PxToCm, 2)
        End If

        Return p
    End Function

    Friend Function GetOriginalPos(transformedPoint As Point, Optional relativeTo As FrameworkElement = Nothing) As Point
        Dim p As Point = transformedPoint

        ' Reverse TransformToVisual
        If relativeTo Is Nothing Then
            p = Pnl.FocusRectangle.TransformToVisual(Canv).Inverse.Transform(p)
        Else
            p = Pnl.FocusRectangle.TransformToVisual(relativeTo).Inverse.Transform(p)
        End If

        ' Reverse RenderTransform
        Dim rt = DesignerItem.RenderTransform
        If rt IsNot Nothing Then p = rt.Inverse.Transform(p)

        ' Reverse LayoutTransform
        Dim Lt = Pnl.FocusRectangle.LayoutTransform
        If Lt IsNot Nothing Then p = Lt.Inverse.Transform(p)

        Return p
    End Function


    Private Sub Diagram_MouseLeave(sender As Object, e As MouseEventArgs) Handles Diagram.MouseLeave
        If DraggingDiagram Then Diagram_PreviewMouseLeftButtonUp(Nothing, Nothing)
        Mouse.OverrideCursor = Nothing
    End Sub

    Private Sub Diagram_GotFocus(sender As Object, e As EventArgs) Handles Diagram.GotFocus
        If Keyboard.Modifiers = ModifierKeys.Control Then
            Pnl.IsSelected = Not Pnl.IsSelected
        ElseIf Pnl.IsSelected = False Then
            Dsn.SelectedIndex = -1
            Pnl.IsSelected = True
        End If
    End Sub


    Friend Sub MoveDiagram(HorizontalChange As Double, VerticalChange As Double, Optional IgnoreRotation As Boolean = False)
        StartMoveUndo()
        DoMoveDiagram(HorizontalChange, VerticalChange, IgnoreRotation)
        ReportMoveUndo()
    End Sub

    Private Sub DoMoveDiagram(HorizontalChange As Double, VerticalChange As Double, Optional IgnoreRotation As Boolean = False)

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
            Pnl.DiagramGroup?.UpdateSelection()
        End If

        For Each D In Dsn.SelectedItems
            Dim Item = Helper.GetListBoxItem(D)
            Canvas.SetLeft(Item, Canvas.GetLeft(Item) + DeltaX)
            Canvas.SetTop(Item, Canvas.GetTop(Item) + DeltaY)
            Helper.GetDiagramPanel(D).DiagramGroup?.UpdateSelection()
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

        Pnl.UpdateTbLocation()
        Helper.UpdateControl(Dsn)

        If Delay > 0 Then Helper.Delay(Delay)

    End Sub

    Private Sub Diagram_PreviewMouseMove(sender As Object, e As MouseEventArgs) Handles Diagram.PreviewMouseMove
        If e.LeftButton = MouseButtonState.Released Then Return

        If DraggingDiagram Then
            Dim p = e.GetPosition(Pnl)
            Dim x = p.X - DraggingStartPoint.X
            Dim y = p.Y - DraggingStartPoint.Y
            DoMoveDiagram(x, y)
            e.Handled = True
        End If
    End Sub

    Private Sub Diagram_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs)
        Diagram.CaptureMouse()
        DraggingDiagram = True

        If DesignerItem.IsSelected Then
            Dsn.SelectedBounds = Dsn.GetSelectionBounds()
            Diagram_GotFocus(Nothing, Nothing)
        Else
            Diagram_GotFocus(Nothing, Nothing)
            Dsn.SelectedBounds = Dsn.GetSelectionBounds()
        End If

        Pnl.UpdateLocationBorder()

        DraggingStartPoint = e.GetPosition(Pnl)

        Pnl.ExitIsSelectedChanged = True
        'TmpIsSelected = Pnl.IsSelected
        DesignerItem.Focus()
        'Pnl.IsSelected = False

        StartMoveUndo()

        e.Handled = True
    End Sub

    Private Sub Diagram_PreviewMouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Diagram.PreviewMouseLeftButtonUp
        Diagram.ReleaseMouseCapture()
        DraggingDiagram = False
        Pnl.ExitIsSelectedChanged = False
        DesignerItem.Focus()
        Dsn.LocationVisibility = Visibility.Collapsed
        ReportMoveUndo()
    End Sub

    Private Sub StartMoveUndo()
        OldPropertyState = New PropertyState(AddressOf AfterRestoreAction, Diagram, Designer.LeftProperty, Designer.TopProperty)
        Unit = New UndoRedoUnit(OldPropertyState)
        For Each D As FrameworkElement In Dsn.SelectedItems
            If D Is Diagram Then Continue For
            Dim Ps As New PropertyState(AddressOf AfterRestoreAction, D, OldPropertyState.Keys.ToArray)
            Unit.Add(Ps)
        Next
    End Sub

    Private Sub ReportMoveUndo()
        If OldPropertyState IsNot Nothing AndAlso OldPropertyState.HasChanges Then
            For Each State As PropertyState In Unit
                State.SetNewValues()
            Next
            Dsn.UndoStack.ReportChanges(Unit)
        End If
    End Sub

#Region "Editor"

    Dim rename As Boolean

    Friend Sub BeginEdit(Optional rename As Boolean = False)
        Me.rename = rename
        Designer.Editing = True
        Dim P = GetLeftTopPoint(InCm:=False)
        Dsn.Editor.Width = Diagram.ActualWidth
        Dsn.Editor.Height = Diagram.ActualHeight
        Dsn.Editor.Text = If(rename, Automation.AutomationProperties.GetName(Diagram), Dsn.GetControlText(Diagram))

        Dim Pup As Popup = Dsn.Editor.Parent
        Pup.PlacementTarget = Diagram
        Pup.Placement = PlacementMode.Relative
        Pup.StaysOpen = True
        Pup.IsOpen = True
        Dsn.Editor.Focus()
        Dsn.Editor.SelectAll()

        AddHandler Dsn.Editor.PreviewKeyDown, AddressOf Editor_PreviewKeyDown
        AddHandler Dsn.Editor.PreviewTextInput, AddressOf Editor_PreviewTextInput
        AddHandler Dsn.Editor.LostFocus, AddressOf Editor_LostFocus
        AddHandler Dsn.PreviewMouseLeftButtonDown, AddressOf Dsn_LeftButtonDown

    End Sub

    Private Sub Dsn_LeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim editor = Helper.GetParent(Of TextBox)(e.OriginalSource)
        If editor IsNot Dsn.Editor Then
            exitLostFocus = True
            EndEdit(True)
            e.Handled = True
            Dsn.Editor.Focus()
            Helper.Dispatcher.BeginInvoke(
                    System.Windows.Threading.DispatcherPriority.Background,
                    Sub() exitLostFocus = False
            )
        End If
    End Sub

    Private Sub Editor_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
        If Not rename Then Return

        Select Case e.Text.ToLower
            Case "a" To "z", "_", "0" To "9"
                ' allowed
            Case Else
                Beep()
                e.Handled = True
        End Select

    End Sub

    Friend Function EndEdit(Commit As Boolean) As Boolean
        Designer.Editing = False

        If Commit Then
            If rename Then
                exitLostFocus = True
                If Not Dsn.SetControlName(Diagram, Dsn.Editor.Text) Then
                    exitLostFocus = False
                    Return False
                End If
                exitLostFocus = False
            Else
                Dsn.SetControlText(Dsn.SelectedIndex, Dsn.Editor.Text)
            End If
        End If

        RemoveHandler Dsn.Editor.PreviewKeyDown, AddressOf Editor_PreviewKeyDown
        RemoveHandler Dsn.Editor.PreviewTextInput, AddressOf Editor_PreviewTextInput
        RemoveHandler Dsn.Editor.LostFocus, AddressOf Editor_LostFocus
        RemoveHandler Dsn.PreviewMouseLeftButtonDown, AddressOf Dsn_LeftButtonDown

        Dim Pup As Popup = Dsn.Editor.Parent
        Pup.IsOpen = False
        DesignerItem.Focus()
        Return True
    End Function

    Dim exitLostFocus As Boolean

    Public Sub Editor_LostFocus(sender As Object, e As RoutedEventArgs)
        If exitLostFocus Then Return

        If Not EndEdit(True) Then
            e.Handled = True
            Helper.Dispatcher.BeginInvoke(
                 Windows.Threading.DispatcherPriority.Background,
                 Sub() Dsn.Editor.Focus()
            )
        End If
    End Sub


    Private Sub Editor_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            If Keyboard.Modifiers = ModifierKeys.Control Then
                If rename Then
                    If Not EndEdit(True) Then
                        e.Handled = True
                        Dsn.Editor.Focus()
                    End If
                Else
                    Dsn.Editor.SelectedText = vbCrLf
                    Dsn.Editor.SelectionLength = 0
                    Dsn.Editor.SelectionStart += 2
                End If

            ElseIf Not EndEdit(True) Then
                e.Handled = True
                Dsn.Editor.Focus()
            End If
        ElseIf e.Key = Key.Escape Then
            EndEdit(False)
        ElseIf e.Key = Key.Space Then
            If rename Then
                Beep()
                e.Handled = True
            End If
        End If
    End Sub

    Friend Sub DiagramTextBlock_TextChanged()
        'Dsn.SetControlText(Diagram, Pnl.DiagramTextBlock.Text)
    End Sub

#End Region


    Private Sub Diagram_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        If e.ClickCount > 1 Then
            Dsn.OnDiagramDoubleClick(Diagram)
        End If
    End Sub

    Private Sub Diagram_PreviewMouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Diagram.PreviewMouseRightButtonUp
        Pnl.ContextMenu.IsOpen = True
        e.Handled = True
    End Sub

    Private Sub Diagram_MouseEnter(sender As Object, e As MouseEventArgs) Handles Diagram.MouseEnter
        Mouse.OverrideCursor = Cursors.SizeAll
    End Sub
End Class
