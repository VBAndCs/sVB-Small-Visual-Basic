﻿Imports System
Imports System.Collections.Specialized
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Threading

Namespace Microsoft.SmallBasic.Shell
    Public Class MdiViewsControl
        'Inherits ItemsControl

        Public Event ActiveDocumentChanged()

        Private isDragOn As Boolean
        Private originPosition As Point
        Private viewRelativePosition As Point
        Private originTop As Double
        Private originLeft As Double
        Private resizeStartPoint As Point
        Private resizeInProgress As Boolean
        Private originalWidth As Double
        Private originalHeight As Double
        Private originalLeft As Double
        Private originalTop As Double
        Private lastXOffset As Double
        Private lastYOffset As Double
        Private lastOffsetForXOffset As Double
        Private selectionChanging As Boolean
        Private lastZIndex As Integer

        Public Property SelectedItem As MdiView

        Public Event RequestItemClose As EventHandler(Of RequestCloseEventArgs)
        Public Shared CloseViewCommand As RoutedUICommand

        Shared Sub New()
            CloseViewCommand = New RoutedUICommand("Close View", "CloseView", GetType(MdiViewsControl))
            CloseViewCommand.InputGestures.Add(New KeyGesture(Key.F4, ModifierKeys.Control))
        End Sub

        Public Sub New()
            Me.InitializeComponent()
            CommandBindings.Add(New CommandBinding(CloseViewCommand, AddressOf CloseView))
            App.GlobalDomain.AddComponent(Me)
            App.GlobalDomain.Bind()
        End Sub

        Protected Overrides Function IsItemItsOwnContainerOverride(ByVal item As Object) As Boolean
            Return TypeOf item Is MdiView
        End Function

        Protected Overrides Function GetContainerForItemOverride() As DependencyObject
            Return New MdiView()
        End Function

        Protected Overrides Sub PrepareContainerForItemOverride(ByVal element As DependencyObject, ByVal item As Object)
            MyBase.PrepareContainerForItemOverride(element, item)
            Dim mdiView = TryCast(element, MdiView)
            If mdiView Is Nothing Then Return

            Dim __ = mdiView.Width
            __ = mdiView.Height
            Dim left = Canvas.GetLeft(mdiView)
            Dim top = Canvas.GetTop(mdiView)

            If Double.IsNaN(left) OrElse Double.IsNaN(top) Then
                left = lastXOffset
                top = lastYOffset
                lastXOffset += 32.0
                lastYOffset += 32.0

                If lastXOffset > ActualWidth OrElse lastYOffset > ActualHeight Then
                    lastOffsetForXOffset += 32.0
                    lastXOffset = lastOffsetForXOffset
                    lastYOffset = 0.0
                End If
            End If

            Canvas.SetLeft(mdiView, left)
            Canvas.SetTop(mdiView, top)
        End Sub

        Private Sub OnBeginDrag(sender As Object, e As MouseEventArgs)
            If e.LeftButton = MouseButtonState.Pressed Then
                isDragOn = True
                Dim uIElement = TryCast(sender, UIElement)
                Dim mdiView = FindViewContainingTemplateItem(uIElement)
                originTop = Canvas.GetTop(mdiView)
                originLeft = Canvas.GetLeft(mdiView)
                originPosition = e.GetPosition(Me)
                viewRelativePosition = e.GetPosition(mdiView)
                uIElement.CaptureMouse()
            End If
        End Sub

        Private Sub OnDrag(sender As Object, e As MouseEventArgs)
            If isDragOn Then
                Dim element As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))
                Dim position = e.GetPosition(Me)

                If position.X < ActualWidth AndAlso position.Y < ActualHeight AndAlso position.X > 0.0 AndAlso position.Y > 0.0 Then
                    Canvas.SetLeft(element, originLeft + (position.X - originPosition.X))
                    Canvas.SetTop(element, originTop + (position.Y - originPosition.Y))
                End If
            End If
        End Sub

        Private Sub OnEndDrag(sender As Object, e As MouseEventArgs)
            If isDragOn Then
                isDragOn = False
                Dim uIElement As UIElement = TryCast(sender, UIElement)
                uIElement.ReleaseMouseCapture()
            End If
        End Sub

        Private Sub OnInitResize(sender As Object, e As MouseEventArgs)
            Dim mdiView As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))
            resizeStartPoint = e.GetPosition(mdiView)
            resizeInProgress = True
            originalWidth = mdiView.ActualWidth
            mdiView.Width = originalWidth
            originalHeight = mdiView.ActualHeight
            mdiView.Height = originalHeight
            originalLeft = Canvas.GetLeft(mdiView)
            originalTop = Canvas.GetTop(mdiView)
            Mouse.Capture(CType(sender, IInputElement))
            mdiView.ResetOldWidthAndHeight()
        End Sub

        Private Sub OnEndResize(sender As Object, e As MouseEventArgs)
            Mouse.Capture(Nothing)
            resizeInProgress = False
        End Sub

        Private Sub OnResizeRightEdge(sender As Object, e As MouseEventArgs)
            If resizeInProgress Then
                Dim mdiView As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))
                Dim position = e.GetPosition(mdiView)
                mdiView.Width = Math.Max(mdiView.MinWidth, originalWidth + (position.X - resizeStartPoint.X))
            End If
        End Sub

        Private Sub OnResizeLeftEdge(sender As Object, e As MouseEventArgs)
            If resizeInProgress Then
                Dim mdiView As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))
                Dim position = e.GetPosition(mdiView)

                If Not mdiView.ActualWidth - (position.X - resizeStartPoint.X) < mdiView.MinWidth Then
                    Canvas.SetLeft(mdiView, Canvas.GetLeft(mdiView) + position.X - resizeStartPoint.X)
                    mdiView.Width = Math.Max(mdiView.MinWidth, originalWidth + (originalLeft - Canvas.GetLeft(mdiView)))
                End If
            End If
        End Sub

        Private Sub OnResizeBottomEdge(sender As Object, e As MouseEventArgs)
            If resizeInProgress Then
                Dim mdiView As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))
                Dim position = e.GetPosition(mdiView)
                mdiView.Height = Math.Max(mdiView.MinHeight, originalHeight + (position.Y - resizeStartPoint.Y))
            End If
        End Sub

        Private Sub OnResizeTopEdge(sender As Object, e As MouseEventArgs)
            If resizeInProgress Then
                Dim mdiView As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))
                Dim position = e.GetPosition(mdiView)

                If Not mdiView.ActualHeight - (position.Y - resizeStartPoint.Y) < mdiView.MinHeight Then
                    Canvas.SetTop(mdiView, Canvas.GetTop(mdiView) + position.Y - resizeStartPoint.Y)
                    mdiView.Height = Math.Max(mdiView.MinHeight, originalHeight + (originalTop - Canvas.GetTop(mdiView)))
                End If
            End If
        End Sub

        Private Sub OnResizeTopLeftCorner(sender As Object, e As MouseEventArgs)
            If resizeInProgress Then
                OnResizeTopEdge(sender, e)
                OnResizeLeftEdge(sender, e)
            End If
        End Sub

        Private Sub OnResizeTopRightCorner(sender As Object, e As MouseEventArgs)
            If resizeInProgress Then
                OnResizeTopEdge(sender, e)
                OnResizeRightEdge(sender, e)
            End If
        End Sub

        Private Sub OnResizeBottomLeftCorner(sender As Object, e As MouseEventArgs)
            If resizeInProgress Then
                OnResizeBottomEdge(sender, e)
                OnResizeLeftEdge(sender, e)
            End If
        End Sub

        Private Sub OnResizeBottomRightCorner(sender As Object, e As MouseEventArgs)
            If resizeInProgress Then
                OnResizeBottomEdge(sender, e)
                OnResizeRightEdge(sender, e)
            End If
        End Sub

        Private Sub CloseView(sender As Object, e As ExecutedRoutedEventArgs)
            Dim mdiView As MdiView = FindViewContainingTemplateItem(TryCast(e.OriginalSource, UIElement))

            If mdiView IsNot Nothing AndAlso RequestItemCloseEvent IsNot Nothing Then
                RaiseEvent RequestItemClose(Me, New RequestCloseEventArgs(mdiView))
            End If
        End Sub

        Private Function FindViewContainingTemplateItem(ByVal templateItem As UIElement) As MdiView
            Dim dependencyObject As DependencyObject = templateItem

            While dependencyObject IsNot Nothing

                If TypeOf dependencyObject Is MdiView Then
                    Return CType(dependencyObject, MdiView)
                End If

                dependencyObject = VisualTreeHelper.GetParent(dependencyObject)
            End While

            Return Nothing
        End Function

        Dim LastselectedView As MdiView

        Friend Sub ChangeSelection(selectedView As MdiView)
            If selectedView Is LastselectedView Then
                Return
            Else
                LastselectedView = selectedView
            End If

            If selectedView Is Nothing Then
                Dim num = 0
                _SelectedItem = Nothing

                For Each item In Items
                    Dim mdiView = TryCast(item, MdiView)
                    Dim zIndex = Panel.GetZIndex(mdiView)

                    If zIndex > num Then
                        num = zIndex
                        selectedView = mdiView
                    End If
                Next

            ElseIf selectedView IsNot _SelectedItem Then
                For Each item In Items
                    Dim mdiView = TryCast(item, MdiView)
                    mdiView.IsSelected = mdiView Is selectedView
                Next

                Panel.SetZIndex(selectedView, lastZIndex)
                lastZIndex += 1
                _SelectedItem = selectedView
                selectedView.IsSelected = True

                selectedView.Dispatcher.Invoke(
                    DispatcherPriority.Render,
                   Sub()
                       KeepFocus = True
                       selectedView.Document.Focus()
                       KeepFocus = False
                   End Sub)
            End If

            RaiseEvent ActiveDocumentChanged()
        End Sub

        Protected Overrides Sub OnItemsChanged(ByVal e As NotifyCollectionChangedEventArgs)
            If Not selectionChanging Then
                If e.OldItems IsNot Nothing Then
                    For Each oldItem In e.OldItems
                        If CType(oldItem, MdiView).IsSelected Then
                            ChangeSelection(Nothing)
                        End If
                    Next
                End If

                If e.NewItems IsNot Nothing Then
                    Dim mdiView2 = TryCast(e.NewItems(e.NewItems.Count - 1), MdiView)

                    If mdiView2 IsNot Nothing Then
                        ChangeSelection(mdiView2)
                    End If
                End If
            End If

            MyBase.OnItemsChanged(e)
        End Sub

        Dim KeepFocus As Boolean
        Private Sub OnFocusWithinChanged(ByVal sender As Object, ByVal e As DependencyPropertyChangedEventArgs)
            If KeepFocus Then Return

            Dim selectedView As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))
            ChangeSelection(selectedView)
        End Sub

        Private Sub OnMouseDownInView(ByVal sender As Object, ByVal e As MouseEventArgs)
            Dim uIElement = CType(sender, UIElement)
            Dim selectedView = FindViewContainingTemplateItem(uIElement)
            uIElement.Focus()
            ChangeSelection(selectedView)
        End Sub

        Private Sub ControlNames_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            Dim cmb = CType(sender, ComboBox)
            Dim controlName = CStr(cmb.SelectedItem)
            Dim selectedView As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))

            If controlName = "" Then
                SelectHandlers(selectedView, "")
                Return
            End If

            Dim events = selectedView.Document.ControlEvents
            events.Clear()

            If cmb.SelectedIndex = 0 Then '  Global
                For Each sb In selectedView.Document.GlobalSubs
                    events.Add(sb)
                Next
            Else
                Dim typeName = selectedView.Document.ControlsInfo(controlName.ToLower())
                For Each ev In WinForms.PreCompiler.GetEvents(typeName)
                    events.Add(ev)
                Next

                SelectHandlers(selectedView, controlName)
            End If
        End Sub

        Private Sub SetItemsBold(cmb As ComboBox, events() As String)
            Dim isGlobal = events Is Nothing
            For i = 0 To cmb.Items.Count - 1
                Dim item = CType(cmb.ItemContainerGenerator.ContainerFromIndex(i), ComboBoxItem)
                If item Is Nothing Then Return
                item.FontWeight = If((isGlobal AndAlso i > 0) OrElse (Not isGlobal AndAlso events.Contains(cmb.Items(i))), FontWeights.Bold, FontWeights.Normal)
            Next
        End Sub

        Dim JustFocus As Boolean

        Private Sub EventNames_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If JustFocus Then Return

            Dim selectedView As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))

            If selectedView.FreezeCmbEvents Then Return

            Dim cmb = CType(sender, ComboBox)
            Dim eventName = CStr(cmb.SelectedItem)
            If eventName = "" Then Return

            selectedView.Document.AddEventHandler(selectedView.CmbControlNames.SelectedItem, eventName)
        End Sub

        Private Sub CmbEventNames_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
            Dim selectedView As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))
            If selectedView.CmbControlNames.SelectedIndex = 0 Then Return

            Dim cmb = CType(sender, ComboBox)
            Dim c = e.Text.ToLower()(0)
            Dim items = cmb.Items

            Dim st As Integer = cmb.SelectedIndex
            For i = 0 To items.Count - 1
                Dim item = CType(cmb.ItemContainerGenerator.ContainerFromIndex(i), ComboBoxItem)
                If item.IsHighlighted Then
                    st = i
                    Exit For
                End If
            Next

            For i = st + 1 To items.Count - 1
                Dim eventName = CStr(items(i)).ToLower()
                If eventName(2) = c Then
                    HighlightItem(cmb, i)
                    e.Handled = True
                    Return
                End If
            Next

            For i = 0 To st
                Dim eventName = CStr(items(i)).ToLower()
                If eventName(2) = c Then
                    HighlightItem(cmb, i)
                    e.Handled = True
                    Return
                End If
            Next
        End Sub

        Dim _highLightedItem As ComboBoxItem

        Private Sub HighlightItem(cmb As ComboBox, index As Integer)
            If index = -1 Then Return

            JustFocus = True
            cmb.SelectedIndex = -1
            cmb.SelectedIndex = index
            JustFocus = False
            _highLightedItem = CType(cmb.ItemContainerGenerator.ContainerFromIndex(index), ComboBoxItem)
        End Sub

        Private Sub CmbEventNames_KeyDown(sender As Object, e As KeyEventArgs)
            If e.Key = Key.Enter Then
                CmbEventNames_PreviewMouseLeftButtonDown(sender, e)
            End If
        End Sub

        Private Sub CmbEventNames_PreviewMouseLeftButtonDown(sender As Object, e As RoutedEventArgs)
            If _highLightedItem Is Nothing Then Return

            Dim cmb = CType(sender, ComboBox)
            If e.OriginalSource IsNot _highLightedItem AndAlso
                GetParent(Of ComboBoxItem)(e.OriginalSource) IsNot _highLightedItem Then Return

            _highLightedItem = Nothing
            EventNames_SelectionChanged(sender, Nothing)

            Dim selectedView As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))
            selectedView.Document.Focus()
            e.Handled = True
        End Sub

        Private Sub CmbEventNames_DropDownOpened(sender As Object, e As EventArgs)
            DiagramHelper.RunAfter.Start(1,
                Sub()
                    Dim view As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))
                    Dim cmbControls = view.CmbControlNames
                    Dim CmbEvents = view.CmbEventNames

                    HighlightItem(sender, CmbEvents.SelectedIndex)
                    If cmbControls.SelectedIndex = 0 Then ' Global
                        SetItemsBold(CmbEvents, Nothing)
                    Else
                        SelectHandlers(view, cmbControls.SelectedItem)
                    End If
                End Sub)
        End Sub

        Sub SelectHandlers(selectedView As MdiView, controlName As String)
            If controlName = "" Then
                SetItemsBold(selectedView.CmbEventNames, {""})
                Return
            End If

            Dim handlers = From h In selectedView.Document.EventHandlers
                           Where h.Value.ControlName = controlName
                           Order By h.Value.EventName
                           Select h.Value.EventName

            If handlers.Any Then
                If selectedView.CmbEventNames.SelectedIndex = -1 Then
                    Dim h = controlName & "_" & handlers.First
                    If selectedView.Document.FindEventHandler(h) = -1 Then
                        selectedView.Document.EventHandlers.Remove(h)
                        SelectHandlers(selectedView, controlName)
                        Return
                    Else
                        selectedView.CmbEventNames.SelectedItem = handlers.First
                    End If
                End If
                SetItemsBold(selectedView.CmbEventNames, handlers.ToArray())
            End If
        End Sub
    End Class
End Namespace
