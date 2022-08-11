Imports System.Collections.ObjectModel
Imports System.Runtime.InteropServices

Namespace Microsoft.Nautilus.Text.Editor
    Friend NotInheritable Class DefaultViewScroller
        Implements IViewScroller

        Private _textView As ITextView

        Public Sub New(textView As ITextView)
            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If
            _textView = textView
        End Sub

        Public Sub ScrollViewportVerticallyByPixels(pixelsToScroll As Double) Implements IViewScroller.ScrollViewportVerticallyByPixels
            If Double.IsNaN(pixelsToScroll) Then
                Throw New ArgumentOutOfRangeException("pixelsToScroll")
            End If

            Dim formattedTextLines1 As IFormattedTextLineCollection = _textView.FormattedTextLines
            If formattedTextLines1.Count > 0 Then
                Dim firstFullyVisibleLine1 As ITextLine = formattedTextLines1.FirstFullyVisibleLine
                _textView.DisplayTextLineContainingCharacter(firstFullyVisibleLine1.LineSpan.Start, firstFullyVisibleLine1.Top + pixelsToScroll, ViewRelativePosition.Top)
            End If
        End Sub

        Public Sub ScrollViewportVerticallyByLine(direction As ScrollDirection) Implements IViewScroller.ScrollViewportVerticallyByLine
            If direction < ScrollDirection.Up OrElse direction > ScrollDirection.Down Then
                Throw New ArgumentOutOfRangeException("direction")
            End If

            Dim formattedTextLines1 As IFormattedTextLineCollection = _textView.FormattedTextLines
            If formattedTextLines1.Count <= 0 Then
                Return
            End If

            Dim indexOfTextLine As Integer = formattedTextLines1.GetIndexOfTextLine(formattedTextLines1.FirstFullyVisibleLine)
            If direction = ScrollDirection.Up Then
                If indexOfTextLine > 0 Then
                    Dim textLine As ITextLine = formattedTextLines1(indexOfTextLine - 1)
                    _textView.DisplayTextLineContainingCharacter(textLine.LineSpan.Start, 0.0, ViewRelativePosition.Top)
                End If
            ElseIf indexOfTextLine < formattedTextLines1.Count - 1 Then
                Dim textLine2 As ITextLine = formattedTextLines1(indexOfTextLine + 1)
                _textView.DisplayTextLineContainingCharacter(textLine2.LineSpan.Start, 0.0, ViewRelativePosition.Top)
            End If
        End Sub

        Public Sub ScrollViewportVerticallyByPage(direction As ScrollDirection) Implements IViewScroller.ScrollViewportVerticallyByPage
            If direction < ScrollDirection.Up OrElse direction > ScrollDirection.Down Then
                Throw New ArgumentOutOfRangeException("direction")
            End If

            Dim formattedTextLines1 As IFormattedTextLineCollection = _textView.FormattedTextLines
            If formattedTextLines1.Count <= 0 Then
                Return
            End If

            If direction = ScrollDirection.Up Then
                Dim index As Integer = Math.Max(0, formattedTextLines1.GetIndexOfTextLine(formattedTextLines1.FirstFullyVisibleLine) - 1)
                Dim textLine As ITextLine = formattedTextLines1(index)
                _textView.DisplayTextLineContainingCharacter(textLine.LineSpan.Start, 0.0, ViewRelativePosition.Bottom)
                Dim num As Integer = formattedTextLines1.GetIndexOfTextLine(textLine) + 1
                If num < formattedTextLines1.Count Then
                    Dim firstFullyVisibleLine1 As ITextLine = formattedTextLines1.FirstFullyVisibleLine
                    If firstFullyVisibleLine1.Top > 0.0 AndAlso formattedTextLines1(num).Bottom - _textView.ViewportHeight > firstFullyVisibleLine1.Top Then
                        _textView.DisplayTextLineContainingCharacter(firstFullyVisibleLine1.LineSpan.Start, 0.0, ViewRelativePosition.Top)
                    End If
                End If
            Else
                Dim index2 As Integer = Math.Min(formattedTextLines1.Count - 1, formattedTextLines1.GetIndexOfTextLine(formattedTextLines1.LastFullyVisibleLine) + 1)
                Dim textLine As ITextLine = formattedTextLines1(index2)
                _textView.DisplayTextLineContainingCharacter(textLine.LineSpan.Start, 0.0, ViewRelativePosition.Top)
            End If
        End Sub

        Public Sub ScrollViewportHorizontallyByPixels(pixelsToScroll As Double) Implements IViewScroller.ScrollViewportHorizontallyByPixels
            If Double.IsNaN(pixelsToScroll) Then
                Throw New ArgumentOutOfRangeException("pixelsToScroll")
            End If
            _textView.ViewportLeft += pixelsToScroll
        End Sub

        Public Function EnsureSpanVisible(span As Span, horizontalPadding As Double, verticalPadding As Double) As Boolean Implements IViewScroller.EnsureSpanVisible
            If Double.IsNaN(horizontalPadding) OrElse horizontalPadding < 0.0 Then
                Throw New ArgumentOutOfRangeException("horizontalPadding")
            End If

            If Double.IsNaN(verticalPadding) OrElse verticalPadding < 0.0 Then
                Throw New ArgumentOutOfRangeException("verticalPadding")
            End If

            If span.End > _textView.TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            Dim formattedTextLines = _textView.FormattedTextLines
            If formattedTextLines.Count = 0 Then
                Return False
            End If

            Dim startTextLine = formattedTextLines.GetTextLineContainingPosition(span.Start)
            Dim endTextLine = formattedTextLines.GetTextLineContainingPosition(span.End)
            Dim startIsNotVisible = span.Start < formattedTextLines(0).LineSpan.Start
            Dim textLine = formattedTextLines(formattedTextLines.Count - 1)
            Dim endIsNotVisible = span.End >= textLine.LineSpan.End AndAlso Not textLine.ContainsPosition(span.End)

            If startIsNotVisible Then
                If endIsNotVisible Then Return False
                _textView.DisplayTextLineContainingCharacter(span.Start, verticalPadding, ViewRelativePosition.Top)
            ElseIf endIsNotVisible Then
                _textView.DisplayTextLineContainingCharacter(span.End, verticalPadding, ViewRelativePosition.Bottom)
            ElseIf startTextLine?.Top < verticalPadding Then
                _textView.DisplayTextLineContainingCharacter(span.Start, verticalPadding, ViewRelativePosition.Top)
            ElseIf endTextLine?.Bottom + verticalPadding > _textView.ViewportHeight Then
                _textView.DisplayTextLineContainingCharacter(span.End, verticalPadding, ViewRelativePosition.Bottom)
            End If


            Dim result As Boolean = False
            If startTextLine Is Nothing OrElse startTextLine.IsDisposed Then Return False
            If endTextLine Is Nothing OrElse endTextLine.IsDisposed Then Return False

            result = startTextLine.Top >= verticalPadding AndAlso endTextLine.Bottom + verticalPadding <= _textView.ViewportHeight
            Dim leftEdge, rightEdge As Double
            ComputeLeftAndRightEdgesOfSpanText(span, formattedTextLines, startTextLine, endTextLine, leftEdge, rightEdge)

            If leftEdge = Double.MaxValue Then
                result = result AndAlso horizontalPadding = 0.0
            Else
                HandleNoWordWrap(horizontalPadding, leftEdge, rightEdge)
                result = result AndAlso leftEdge >= horizontalPadding AndAlso _textView.ViewportWidth - rightEdge >= horizontalPadding
            End If

            Return result
        End Function

        Private Sub HandleNoWordWrap(horizontalPadding As Double, ByRef leftEdge As Double, ByRef rightEdge As Double)
            If (_textView.WordWrapStyle And WordWrapStyles.WordWrap) = WordWrapStyles.WordWrap Then
                Return
            End If

            Dim viewportLeft1 As Double = _textView.ViewportLeft
            Dim num As Double = viewportLeft1 - (leftEdge - horizontalPadding)
            Dim num2 As Double = rightEdge + horizontalPadding - (viewportLeft1 + _textView.ViewportWidth)

            If num > 0.0 Then
                If num2 <= 0.0 Then
                    _textView.ViewportLeft = viewportLeft1 - num
                End If
            ElseIf num2 > 0.0 Then
                _textView.ViewportLeft = viewportLeft1 + num2
            End If

            Dim num3 As Double = _textView.ViewportLeft - viewportLeft1
            leftEdge += num3
            rightEdge += num3
        End Sub

        Private Sub ComputeLeftAndRightEdgesOfSpanText(span1 As Span, textLines As IFormattedTextLineCollection, startLine As ITextLine, endLine As ITextLine, <Out> ByRef leftEdge As Double, <Out> ByRef rightEdge As Double)
            leftEdge = Double.MaxValue
            rightEdge = Double.MinValue
            If startLine Is endLine Then
                ExpandBounds(startLine, span1, leftEdge, rightEdge)
                Return
            End If

            ExpandBounds(startLine, New Span(span1.Start, startLine.LineSpan.End - span1.Start), leftEdge, rightEdge)
            ExpandBounds(endLine, New Span(endLine.LineSpan.Start, span1.End - endLine.LineSpan.Start), leftEdge, rightEdge)
            Dim indexOfTextLine As Integer = textLines.GetIndexOfTextLine(startLine)
            Dim indexOfTextLine2 As Integer = textLines.GetIndexOfTextLine(endLine)
            For i As Integer = indexOfTextLine + 1 To indexOfTextLine2 - 1
                Dim textLine As ITextLine = textLines(i)
                ExpandBounds(textLine, New Span(textLine.LineSpan.Start, textLine.LineSpan.Length), leftEdge, rightEdge)
            Next
        End Sub

        Private Sub ExpandBounds(line As ITextLine, span1 As Span, ByRef leftEdge As Double, ByRef rightEdge As Double)
            Dim textBounds1 As ReadOnlyCollection(Of TextBounds) = line.GetTextBounds(span1)
            For Each item As TextBounds In textBounds1
                If item.Left < leftEdge Then
                    leftEdge = item.Left
                End If
                If item.Right > rightEdge Then
                    rightEdge = item.Right
                End If
            Next
        End Sub
    End Class
End Namespace
