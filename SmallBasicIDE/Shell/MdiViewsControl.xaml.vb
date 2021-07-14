Imports System
Imports System.Collections.Specialized
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
            Dim x = New CommandBinding(CloseViewCommand, AddressOf CloseView)
            CommandBindings.Add(x)
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

            If mdiView Is Nothing Then
                Return
            End If

            Dim _1 = mdiView.Width
            Dim _2 = mdiView.Height
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

        Friend Sub ChangeSelection(ByVal selectedView As MdiView)
            If selectedView Is LastselectedView Then
                Return
            Else
                LastselectedView = selectedView
            End If

            If selectedView Is Nothing Then
                Dim num = 0

                For Each item In Items
                    Dim mdiView = TryCast(item, MdiView)
                    Dim zIndex = Panel.GetZIndex(mdiView)

                    If zIndex > num Then
                        num = zIndex
                        selectedView = mdiView
                    End If
                Next
            End If

            If selectedView Is Nothing Then
                _SelectedItem = Nothing
            ElseIf selectedView IsNot _SelectedItem Then
                For Each item2 In Items
                    Dim mdiView2 = TryCast(item2, MdiView)

                    If mdiView2 IsNot selectedView Then
                        mdiView2.IsSelected = False
                    End If
                Next
                Panel.SetZIndex(selectedView, lastZIndex)
                lastZIndex += 1
                _SelectedItem = selectedView
                selectedView.IsSelected = True
                selectedView.Dispatcher.Invoke(DispatcherPriority.Render, Sub() Exit Sub)

                KeepFocus = True
                CType(selectedView.Document.EditorControl.TextView, Nautilus.Text.Editor.AvalonTextView).Focus()
                KeepFocus = False
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
            If controlName = "" Then
                cmb.SelectedIndex = 0
                Return
            End If

            Dim selectedView As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))
            Dim events = selectedView.Document.ControlEvents
            events.Clear()

            If controlName = "(Global)" Then
                For Each sb In selectedView.Document.GlobalSubs
                    events.Add(sb)
                Next
            Else
                Dim typeName = selectedView.Document.ControlsInfo(controlName.ToLower())
                For Each ev In WinForms.PreCompiler.GetEvents(typeName)
                    events.Add(ev)
                Next
            End If
        End Sub

        Private Sub EventNames_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            Dim selectedView As MdiView = FindViewContainingTemplateItem(TryCast(sender, UIElement))
            If selectedView.FreezeCmbEvents Then Return

            Dim cmb = CType(sender, ComboBox)
            Dim eventName = CStr(cmb.SelectedItem)
            If eventName = "" Then Return

            selectedView.FreezeCmbEvents = True
            selectedView.Document.AddEventHandler(selectedView.CmbControlNames.SelectedItem, eventName)
            selectedView.FreezeCmbEvents = False
        End Sub
    End Class
End Namespace
