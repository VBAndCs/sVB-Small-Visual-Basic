Imports System.Windows.Controls.Primitives
Imports System.Windows.Media.Animation
Imports sb = Microsoft.SmallBasic

Namespace Microsoft.SmallBasic.LanguageService
    Partial Public Class CircularList
        Inherits Selector

        Private _itemContainerMap As Dictionary(Of Integer, CircularListItem) = New Dictionary(Of Integer, CircularListItem)()
        Private _itemsHost As Canvas
        Private _isAnimating As Boolean
        Private _oldSelectionIndex As Integer = -1

        Public Sub New()
            Me.InitializeComponent()
        End Sub

        Public Overrides Sub OnApplyTemplate()
            _itemsHost = CType(Template.FindName("itemsHost", Me), Canvas)
            Dispatcher.BeginInvoke(CType(Sub() ArrangeItems(), Action))
        End Sub

        Protected Overrides Sub OnItemsSourceChanged(
                         oldValue As IEnumerable,
                         newValue As IEnumerable
                    )

            _itemContainerMap.Clear()

            If _itemsHost IsNot Nothing Then
                _itemsHost.Children.Clear()
            End If

            Dim index = 0

            For Each item In newValue
                _itemContainerMap(index) = New CircularListItem(Me, index) With {.Content = item}
                index += 1
            Next

            MyBase.OnItemsSourceChanged(oldValue, newValue)
        End Sub

        Public Sub MoveUp()
            If Not _isAnimating Then
                SelectedIndex = If(SelectedIndex > 0, SelectedIndex - 1, Items.Count - 1)
            End If
        End Sub

        Public Sub MoveDown()
            If Not _isAnimating Then
                SelectedIndex = If(SelectedIndex < Items.Count - 1, SelectedIndex + 1, 0)
            End If
        End Sub

        Protected Overrides Sub OnSelectionChanged(e As SelectionChangedEventArgs)
            If SelectedIndex <> -1 Then
                Dispatcher.BeginInvoke(Sub() ArrangeItems())
                MyBase.OnSelectionChanged(e)
            End If
        End Sub

        Private Sub ArrangeItems()
            Dim _moveDirection = If(_oldSelectionIndex > SelectedIndex,
                 MoveDirection.Down, MoveDirection.Up)

            _oldSelectionIndex = SelectedIndex
            Dim list As New List(Of Integer)()
            Dim top = 0.0

            For i = -3 To 3
                Dim item = ArrangeItem(i, _moveDirection, top)
                list.Add(item)
            Next

            For j = 0 To Items.Count - 1
                If Not list.Contains(j) Then
                    HideItem(j, _moveDirection)
                End If
            Next
        End Sub

        Private Sub HideItem(itemIndex As Integer, moveDirection As MoveDirection)
            Dim value As CircularListItem = Nothing
            _itemContainerMap.TryGetValue(itemIndex, value)

            If value IsNot Nothing AndAlso _itemsHost IsNot Nothing Then
                value.Scale = 0.1
                Dim top = Canvas.GetTop(value)

                If Double.IsNaN(top) Then top = 0.0

                If moveDirection = MoveDirection.Up Then
                    DoubleAnimateProperty(value, Canvas.TopProperty, top, -18.0, 100.0)
                Else
                    DoubleAnimateProperty(value, Canvas.TopProperty, top, ActualHeight + 18.0, 100.0)
                End If

                _itemsHost.Children.Remove(value)
            End If
        End Sub

        Private Function ArrangeItem(
                     relIndex As Integer,
                     moveDirection As MoveDirection,
                     ByRef top As Double
                  ) As Integer

            Dim index = SelectedIndex + relIndex

            If index < 0 Then
                index = Items.Count + index
            ElseIf index > Items.Count - 1 Then
                index -= Items.Count
            End If

            Dim item As CircularListItem = Nothing
            _itemContainerMap.TryGetValue(index, item)

            If item IsNot Nothing AndAlso _itemsHost IsNot Nothing Then
                If Not _itemsHost.Children.Contains(item) Then
                    _itemsHost.Children.Add(item)
                End If

                If relIndex = 0 Then
                    TextBlock.SetFontWeight(item, FontWeights.Bold)
                Else
                    TextBlock.SetFontWeight(item, FontWeights.Normal)
                End If

                AnimateItem(item, relIndex, moveDirection, top)
            End If

            Return index
        End Function

        Private Sub AnimateItem(
                        cli As CircularListItem,
                        relIndex As Integer,
                        moveDirection As MoveDirection,
                        ByRef top As Double
                  )

            Dim r = 1.7 / Math.Pow(1.33, Math.Abs(relIndex))
            Dim cliTop = Canvas.GetTop(cli)

            If Double.IsNaN(cliTop) OrElse cliTop < 0.0 OrElse cliTop > ActualHeight Then
                cliTop = If(moveDirection <> MoveDirection.Down,
                    ActualHeight, (0.0 - cli.Height) * r)
            End If

            Dim cliLeft = Canvas.GetLeft(cli)
            If Double.IsNaN(cliLeft) Then cliLeft = 0.0

            DoubleAnimateProperty(cli, CircularListItem.ScaleProperty, cli.Scale, r, 100.0)
            DoubleAnimateProperty(cli, Canvas.TopProperty, cliTop, top, 100.0)
            DoubleAnimateProperty(cli, Canvas.LeftProperty, cliLeft, 12 - Math.Abs(relIndex) * 3, 100.0)
            top += (cli.Height + 2.0) * r
        End Sub

        Private Sub DoubleAnimateProperty(
                       animatable As IAnimatable,
                       [property] As DependencyProperty,
                       start As Double,
                       [end] As Double,
                       time As Double)

            _isAnimating = True
            Dim doubleAnimation As New DoubleAnimation(
                    start, [end],
                    New Duration(TimeSpan.FromMilliseconds(time))
             )

            AddHandler doubleAnimation.Completed,
                        Sub() _isAnimating = False

            animatable.BeginAnimation([property], doubleAnimation)
        End Sub

        Protected Overrides Sub OnMouseDown(e As MouseButtonEventArgs)
            e.Handled = True
            MyBase.OnMouseDown(e)
        End Sub
    End Class
End Namespace
