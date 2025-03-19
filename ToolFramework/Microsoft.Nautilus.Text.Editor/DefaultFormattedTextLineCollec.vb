Imports System.Collections
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

Namespace Microsoft.Nautilus.Text.Editor

    Public NotInheritable Class DefaultFormattedTextLineCollection(Of T As {Class, ITextLine})
        Implements IFormattedTextLineCollection
        Implements IList(Of ITextLine), ICollection(Of ITextLine), IEnumerable(Of ITextLine), IEnumerable

        Private NotInheritable Class MyEnumerator
            Implements IEnumerator(Of ITextLine), IDisposable, IEnumerator

            Private _currentIndex As Integer = -1
            Private _textLines As IList(Of T)

            Public ReadOnly Property Current As ITextLine Implements IEnumerator(Of ITextLine).Current
                Get
                    If _currentIndex < 0 OrElse _currentIndex >= _textLines.Count Then
                        Throw New InvalidOperationException("Enumerator not at a valid position")
                    End If

                    Return _textLines(_currentIndex)
                End Get
            End Property

            ReadOnly Property IEnumerator_Current As Object Implements IEnumerator.Current
                Get
                    If _currentIndex < 0 OrElse _currentIndex >= _textLines.Count Then
                        Throw New InvalidOperationException("Enumerator not at a valid position")
                    End If

                    Return _textLines(_currentIndex)
                End Get
            End Property

            Public Sub New(textLines As IList(Of T))
                _textLines = textLines
            End Sub

            Private Sub Dispose() Implements IDisposable.Dispose
            End Sub

            Public Function MoveNext() As Boolean Implements Collections.IEnumerator.MoveNext
                _currentIndex += 1
                Return _currentIndex < _textLines.Count
            End Function

            Public Sub Reset() Implements Collections.IEnumerator.Reset
                _currentIndex = -1
            End Sub
        End Class

        Private _textView As ITextView
        Private _textLines As IList(Of T)

        Default Public Property Item(index As Integer) As ITextLine Implements Collections.Generic.IList(Of ITextLine).Item
            Get
                Return _textLines(index)
            End Get

            Set(value As ITextLine)
                Throw New InvalidOperationException("Attempt to modify a read only collection.")
            End Set
        End Property

        Public ReadOnly Property Count As Integer Implements ICollection(Of ITextLine).Count
            Get
                Return _textLines.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of ITextLine).IsReadOnly
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property FirstFullyVisibleLine As ITextLine Implements IFormattedTextLineCollection.FirstFullyVisibleLine
            Get
                If _textLines.Count = 0 Then
                    Return Nothing
                End If

                Dim index As Integer
                If Not FindTextLineIndexContainingYCoordinate(0.0, index) Then
                    index += 1
                    If index >= _textLines.Count Then
                        Return Nothing
                    End If

                    Dim textLine As ITextLine = _textLines(index)
                    If GetVisibilityStateOfTextLine(textLine) <> 0 Then
                        Return textLine
                    End If

                    Return Nothing
                End If

                Dim textLine2 As ITextLine = _textLines(index)
                If GetVisibilityStateOfTextLine(textLine2) <> VisibilityState.FullyVisible Then
                    index += 1
                    If index < _textLines.Count Then
                        Dim textLine3 As ITextLine = _textLines(index)
                        If GetVisibilityStateOfTextLine(textLine3) = VisibilityState.FullyVisible Then
                            textLine2 = textLine3
                        End If
                    End If
                End If

                Return textLine2
            End Get
        End Property

        Public ReadOnly Property LastFullyVisibleLine As ITextLine Implements IFormattedTextLineCollection.LastFullyVisibleLine
            Get
                If _textLines.Count = 0 Then
                    Return Nothing
                End If

                Dim index As Integer
                If Not FindTextLineIndexContainingYCoordinate(_textView.ViewportHeight, index) Then
                    If index < 0 Then
                        Return Nothing
                    End If

                    Dim textLine As ITextLine = _textLines(index)
                    If GetVisibilityStateOfTextLine(textLine) <> 0 Then
                        Return textLine
                    End If

                    Return Nothing
                End If

                Dim textLine2 As ITextLine = _textLines(index)
                If GetVisibilityStateOfTextLine(textLine2) <> VisibilityState.FullyVisible AndAlso index > 0 Then
                    Dim textLine3 As ITextLine = _textLines(index - 1)
                    If GetVisibilityStateOfTextLine(textLine3) = VisibilityState.FullyVisible Then
                        textLine2 = textLine3
                    End If
                End If

                Return textLine2
            End Get
        End Property

        Public ReadOnly Property FormattedContentWidth As Double Implements IFormattedTextLineCollection.FormattedContentWidth
            Get
                Dim num As Double = 0.0
                For Each textLine As T In _textLines
                    If textLine.Right > num Then
                        num = textLine.Width
                    End If
                Next

                Return num
            End Get
        End Property

        Public ReadOnly Property FormattedContentHeight As Double Implements IFormattedTextLineCollection.FormattedContentHeight
            Get
                If _textLines.Count <> 0 Then
                    Dim val As T = _textLines(_textLines.Count - 1)
                    Dim bottom1 As Double = val.Bottom
                    Dim val2 As T = _textLines(0)
                    Return bottom1 - val2.Top
                End If

                Return 0.0
            End Get
        End Property

        Public Sub New(textView As ITextView, textLines As IList(Of T))
            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            If textLines Is Nothing Then
                Throw New ArgumentNullException("textLines")
            End If

            _textView = textView
            _textLines = textLines
        End Sub

        Public Function IndexOf(item As ITextLine) As Integer Implements Collections.Generic.IList(Of ITextLine).IndexOf
            Dim val As T = TryCast(item, T)
            If val IsNot Nothing OrElse item Is Nothing Then
                Return _textLines.IndexOf(val)
            End If
            Return -1
        End Function

        Private Sub Insert(index As Integer, item As ITextLine) Implements Collections.Generic.IList(Of ITextLine).Insert
            Throw New InvalidOperationException("Attempt to modify a read only collection.")
        End Sub

        Private Sub RemoveAt(index As Integer) Implements Collections.Generic.IList(Of ITextLine).RemoveAt
            Throw New InvalidOperationException("Attempt to modify a read only collection.")
        End Sub

        Private Sub Add(item As ITextLine) Implements Collections.Generic.ICollection(Of ITextLine).Add
            Throw New InvalidOperationException("Attempt to modify a read only collection.")
        End Sub

        Private Sub Clear() Implements Collections.Generic.ICollection(Of ITextLine).Clear
            Throw New InvalidOperationException("Attempt to modify a read only collection.")
        End Sub

        Public Function Contains(item As ITextLine) As Boolean Implements Collections.Generic.ICollection(Of ITextLine).Contains
            Dim val As T = TryCast(item, T)
            If val IsNot Nothing OrElse item Is Nothing Then
                Return _textLines.Contains(val)
            End If
            Return False
        End Function

        Public Sub CopyTo(array As ITextLine(), arrayIndex As Integer) Implements Collections.Generic.ICollection(Of ITextLine).CopyTo
            For i As Integer = 0 To _textLines.Count - 1
                array(i + arrayIndex) = _textLines(i)
            Next
        End Sub

        Private Function Remove(item As ITextLine) As Boolean Implements Collections.Generic.ICollection(Of ITextLine).Remove
            Throw New InvalidOperationException("Attempt to modify a read only collection.")
        End Function

        Public Function GetEnumerator1() As IEnumerator(Of ITextLine) Implements Collections.Generic.IEnumerable(Of ITextLine).GetEnumerator
            Return New MyEnumerator(_textLines)
        End Function

        Private Function GetEnumerator2() As IEnumerator Implements Collections.IEnumerable.GetEnumerator
            Return New MyEnumerator(_textLines)
        End Function

        Public Function FindTextLineIndexContainingPosition(position As Integer, <Out> ByRef index As Integer) As Boolean Implements IFormattedTextLineCollection.FindTextLineIndexContainingPosition
            If position < 0 OrElse position > _textView.TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            Dim num As Integer = _textLines.Count - 1
            If num >= 0 Then
                Dim val As T = _textLines(0)
                If position >= val.LineSpan.Start Then
                    Dim val2 As T = _textLines(num)
                    If position >= val2.LineSpan.Start Then
                        index = num
                        Dim val3 As T = _textLines(num)
                        Return val3.ContainsPosition(position)
                    End If
                    num -= 1
                    Dim num2 As Integer = 0

                    Do
                        index = (num2 + num) \ 2
                        Dim textLine As ITextLine = _textLines(index)

                        If position < textLine.LineSpan.Start Then
                            index -= 1
                            num = index
                            Continue Do
                        End If

                        If position >= textLine.LineSpan.End Then
                            num2 = index + 1
                            Continue Do
                        End If

                        Return True
                    Loop While num2 <= num

                    Return False
                End If
            End If
            index = -1
            Return False
        End Function

        Public Function GetTextLineContainingPosition(position As Integer) As ITextLine Implements IFormattedTextLineCollection.GetTextLineContainingPosition
            If position < 0 OrElse position > _textView.TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("position")
            End If
            Dim index As Integer = 0
            If FindTextLineIndexContainingPosition(position, index) Then
                Return _textLines(index)
            End If
            Return Nothing
        End Function

        Public Function FindTextLineIndexContainingYCoordinate(y As Double, <Out> ByRef index As Integer) As Boolean Implements IFormattedTextLineCollection.FindTextLineIndexContainingYCoordinate
            If Double.IsNaN(y) Then
                Throw New ArgumentOutOfRangeException("y")
            End If

            Dim num As Integer = _textLines.Count - 1
            If num >= 0 Then
                Dim val As T = _textLines(0)
                If Not (y < val.Top) Then
                    Dim val2 As T = _textLines(num)
                    If y >= val2.Top Then
                        index = num
                        Dim val3 As T = _textLines(num)
                        Dim bottom1 As Double = val3.Bottom
                        If Not (y < bottom1) Then
                            If y = bottom1 Then
                                Dim val4 As T = _textLines(num)
                                If val4.LineSpan.End = _textView.TextSnapshot.Length Then
                                    Dim val5 As T = _textLines(num)
                                    Return val5.LineLength = 0
                                End If
                            End If
                            Return False
                        End If
                        Return True
                    End If
                    num -= 1
                    Dim num2 As Integer = 0

                    Do
                        index = (num2 + num) \ 2
                        Dim textLine As ITextLine = _textLines(index)
                        If y < textLine.Top Then
                            index -= 1
                            num = index
                            Continue Do
                        End If

                        Dim val6 As T = _textLines(index)
                        If y >= val6.Bottom Then
                            num2 = index + 1
                            Continue Do
                        End If

                        Return True
                    Loop While num2 <= num

                    Return False
                End If
            End If

            index = -1
            Return False
        End Function

        Public Function GetTextLineContainingYCoordinate(y As Double) As ITextLine Implements IFormattedTextLineCollection.GetTextLineContainingYCoordinate
            If Double.IsNaN(y) Then
                Throw New ArgumentOutOfRangeException("y")
            End If

            Dim index As Integer = 0
            If FindTextLineIndexContainingYCoordinate(y, index) Then
                Return _textLines(index)
            End If
            Return Nothing
        End Function

        Public Function GetIndexOfTextLine(textLine As ITextLine) As Integer Implements IFormattedTextLineCollection.GetIndexOfTextLine
            If textLine Is Nothing Then
                Throw New ArgumentNullException("textLine")
            End If
            If textLine.IsDisposed Then
                Stop
                Return -1
                Throw New ObjectDisposedException("textLine")
            End If

            Dim index As Integer
            If Not FindTextLineIndexContainingPosition(textLine.LineSpan.Start, index) Then
                Throw New ArgumentException("textLine not found within textlines", "textLine")
            End If
            Return index
        End Function

        Public Function GetVisibilityStateOfTextLine(textLine As ITextLine) As VisibilityState Implements IFormattedTextLineCollection.GetVisibilityStateOfTextLine
            If textLine Is Nothing Then
                Throw New ArgumentNullException("textLine")
            End If
            If textLine.IsDisposed Then
                Throw New ObjectDisposedException("textLine")
            End If
            Dim top1 As Double = textLine.Top
            Dim bottom1 As Double = textLine.Bottom
            Dim viewportHeight1 As Double = _textView.ViewportHeight
            If bottom1 < 0.0 OrElse top1 > viewportHeight1 Then
                Return VisibilityState.Hidden
            End If
            If top1 >= 0.0 AndAlso bottom1 <= viewportHeight1 Then
                Return VisibilityState.FullyVisible
            End If
            Return VisibilityState.PartiallyVisible
        End Function
    End Class
End Namespace
