Imports System.Windows.Controls.Primitives
Imports System.Windows.Media.Animation
Imports sb = Microsoft.SmallBasic

Namespace Microsoft.SmallBasic.LanguageService
    Public Partial Class CircularList
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

        Protected Overrides Sub OnItemsSourceChanged(ByVal oldValue As IEnumerable, ByVal newValue As IEnumerable)
            _itemContainerMap.Clear()

            If _itemsHost IsNot Nothing Then
                _itemsHost.Children.Clear()
            End If

            Dim num = 0

            For Each item In newValue
                _itemContainerMap(num) = New CircularListItem(Me, num) With {
                    .Content = item
                }
                num += 1
            Next

            SelectedIndex = 0
            MyBase.OnItemsSourceChanged(oldValue, newValue)
        End Sub

        Public Sub MoveUp()
            If Not _isAnimating Then
                Dim num = SelectedIndex - 1

                If num < 0 Then
                    num = Items.Count - 1
                End If

                SelectedIndex = num
            End If
        End Sub

        Public Sub MoveDown()
            If Not _isAnimating Then
                Dim num = SelectedIndex + 1

                If num > Items.Count - 1 Then
                    num = 0
                End If

                SelectedIndex = num
            End If
        End Sub

        Protected Overrides Sub OnSelectionChanged(ByVal e As SelectionChangedEventArgs)
            If SelectedIndex <> -1 Then
                Dispatcher.BeginInvoke(CType(Sub() ArrangeItems(), Action))
                MyBase.OnSelectionChanged(e)
            End If
        End Sub

        Private Sub ArrangeItems()
            Dim _moveDirection = MoveDirection.Up

            If _oldSelectionIndex > SelectedIndex Then
                _moveDirection = MoveDirection.Down
            End If

            _oldSelectionIndex = SelectedIndex
            Dim list As List(Of Integer) = New List(Of Integer)()
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

        Private Sub HideItem(ByVal itemIndex As Integer, ByVal moveDirection As MoveDirection)
            Dim value As CircularListItem = Nothing
            _itemContainerMap.TryGetValue(itemIndex, value)

            If value IsNot Nothing AndAlso _itemsHost IsNot Nothing Then
                value.Scale = 0.1
                Dim num = Canvas.GetTop(value)

                If Double.IsNaN(num) Then
                    num = 0.0
                End If

                If moveDirection = MoveDirection.Up Then
                    DoubleAnimateProperty(value, Canvas.TopProperty, num, -18.0, 100.0)
                Else
                    DoubleAnimateProperty(value, Canvas.TopProperty, num, ActualHeight + 18.0, 100.0)
                End If

                _itemsHost.Children.Remove(value)
            End If
        End Sub

        Private Function ArrangeItem(ByVal relIndex As Integer, ByVal moveDirection As MoveDirection, ByRef top As Double) As Integer
            Dim num = SelectedIndex + relIndex

            If num < 0 Then
                num = Items.Count + num
            ElseIf num > Items.Count - 1 Then
                num -= Items.Count
            End If

            Dim value As CircularListItem = Nothing
            _itemContainerMap.TryGetValue(num, value)

            If value IsNot Nothing AndAlso _itemsHost IsNot Nothing Then
                If Not _itemsHost.Children.Contains(value) Then
                    _itemsHost.Children.Add(value)
                End If

                If relIndex = 0 Then
                    TextBlock.SetFontWeight(value, FontWeights.Bold)
                Else
                    TextBlock.SetFontWeight(value, FontWeights.Normal)
                End If

                AnimateItem(value, relIndex, moveDirection, top)
            End If

            Return num
        End Function

        Private Sub AnimateItem(ByVal cli As CircularListItem, ByVal relIndex As Integer, ByVal moveDirection As MoveDirection, ByRef top As Double)
            Dim num = 1.7 / Math.Pow(1.33, Math.Abs(relIndex))
            Dim num2 = Canvas.GetTop(cli)

            If Double.IsNaN(num2) OrElse num2 < 0.0 OrElse num2 > ActualHeight Then
                num2 = If(moveDirection <> MoveDirection.Down, ActualHeight, (0.0 - cli.Height) * num)
            End If

            Dim num3 = Canvas.GetLeft(cli)

            If Double.IsNaN(num3) Then
                num3 = 0.0
            End If

            DoubleAnimateProperty(cli, CircularListItem.ScaleProperty, cli.Scale, num, 100.0)
            DoubleAnimateProperty(cli, Canvas.TopProperty, num2, top, 100.0)
            DoubleAnimateProperty(cli, Canvas.LeftProperty, num3, 12 - Math.Abs(relIndex) * 3, 100.0)
            top += (cli.Height + 2.0) * num
        End Sub

        Private Sub DoubleAnimateProperty(ByVal animatable As IAnimatable, ByVal [property] As DependencyProperty, ByVal start As Double, ByVal [end] As Double, ByVal timespan As Double)
            _isAnimating = True
            Dim doubleAnimation As DoubleAnimation = New DoubleAnimation(start, [end], New Duration(System.TimeSpan.FromMilliseconds(timespan)))
            AddHandler doubleAnimation.Completed, Sub() _isAnimating = False
            animatable.BeginAnimation([property], doubleAnimation)
        End Sub

        Protected Overrides Sub OnMouseDown(ByVal e As MouseButtonEventArgs)
            e.Handled = True
            MyBase.OnMouseDown(e)
        End Sub
    End Class
End Namespace
