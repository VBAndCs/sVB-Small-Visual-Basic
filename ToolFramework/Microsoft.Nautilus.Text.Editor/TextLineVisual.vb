Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.TextFormatting

Namespace Microsoft.Nautilus.Text.Editor

    Friend NotInheritable Class TextLineVisual
        Inherits DrawingVisual
        Implements ITextLine, IDisposable

        Private Const _newLineWidth As Double = 10.0
        Private Const _lineHeightPadding As Double = 2.0
        Private _textView As ITextView
        Private _span As ITextSpan
        Private _sourceIndex As Integer
        Private _textLines As IList(Of TextLine)
        Private _textLineDistances As List(Of Double)
        Private _textLineStartIndices As List(Of Integer)
        Private _newLineLength As Integer
        Private _virtualCharacterPositions As List(Of Integer)
        Private _lineContainsBidi As Boolean?
        Private _horizontalOffset As Double
        Private _verticalOffset As Double
        Private _height As Double
        Private _width As Double
        Private _extent As Double
        Private _baseline As Double
        Private _overhangAfter As Double
        Private _transform As TranslateTransform
        Private _textElementSpans As ArrayList

        Public ReadOnly Property IsDisposed As Boolean Implements ITextLine.IsDisposed
            Get
                Return _textLines Is Nothing
            End Get
        End Property

        Public ReadOnly Property Baseline As Double Implements ITextLine.Baseline
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return _baseline
            End Get
        End Property

        Public ReadOnly Property Extent As Double Implements ITextLine.Extent
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return _extent
            End Get
        End Property

        Public ReadOnly Property OverhangAfter As Double Implements ITextLine.OverhangAfter
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return _overhangAfter
            End Get
        End Property

        Public ReadOnly Property OverhangLeading As Double Implements ITextLine.OverhangLeading
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return _textLines(0).OverhangLeading
            End Get
        End Property

        Public ReadOnly Property OverhangTrailing As Double Implements ITextLine.OverhangTrailing
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return _textLines(_textLines.Count - 1).OverhangTrailing
            End Get
        End Property

        Public ReadOnly Property LineSpan As Span Implements ITextLine.LineSpan
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return _span.GetSpan(_textView.TextSnapshot)
            End Get
        End Property

        Public ReadOnly Property NewlineLength As Integer Implements ITextLine.NewlineLength
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return _newLineLength
            End Get
        End Property

        Public Property Left As Double Implements ITextLine.Left
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return _horizontalOffset
            End Get

            Set(value As Double)
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                _horizontalOffset = value
                _transform.X = value
            End Set
        End Property

        Public Property Top As Double Implements ITextLine.Top
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return _verticalOffset
            End Get

            Set(value As Double)
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                _verticalOffset = value
                _transform.Y = value
            End Set
        End Property

        Public ReadOnly Property Height As Double Implements ITextLine.Height
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return _height
            End Get
        End Property

        Public ReadOnly Property Bottom As Double Implements ITextLine.Bottom
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return Top + Height
            End Get
        End Property

        Public ReadOnly Property Right As Double Implements ITextLine.Right
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return Left + Width
            End Get
        End Property

        Public ReadOnly Property Width As Double Implements ITextLine.Width
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return _width
            End Get
        End Property

        Private ReadOnly Property LineContainsBidi As Boolean
            Get
                If Not _lineContainsBidi.HasValue Then
                    _lineContainsBidi = False
                    For Each textLine1 As TextLine In _textLines
                        For Each indexedGlyphRun As IndexedGlyphRun In textLine1.GetIndexedGlyphRuns()
                            If (CUInt(indexedGlyphRun.GlyphRun.BidiLevel) And (If(True, 1UL, 0UL))) <> 0 Then
                                _lineContainsBidi = True
                                Exit For
                            End If
                        Next
                    Next
                End If

                Return _lineContainsBidi.Value
            End Get
        End Property

        Public Sub Dispose() Implements IDisposable.Dispose
            If Not IsDisposed Then
                For i As Integer = 0 To _textLines.Count - 1
                    _textLines(i).Dispose()
                Next
                _textLines = Nothing
            End If
        End Sub

        Public Function MoveCaretToLocation(horizontalPosition As Double) As ICaretPosition Implements ITextLine.MoveCaretToLocation
            If IsDisposed Then
                Throw New ObjectDisposedException("TextLineVisual")
            End If

            If Double.IsNaN(horizontalPosition) Then
                Throw New ArgumentOutOfRangeException("horizontalPosition")
            End If

            Dim num As Double = horizontalPosition - Left
            If num < 0.0 Then
                Return _textView.Caret.MoveTo(LineSpan.Start, CaretPlacement.LeadingEdgeOfCharacter)
            End If

            If NewlineLength = 0 Then
                If num >= Width Then
                    Return _textView.Caret.MoveTo(LineSpan.End, CaretPlacement.LeadingEdgeOfCharacter)
                End If
            ElseIf num >= Width - 10.0 Then
                Return _textView.Caret.MoveTo(LineSpan.End - NewlineLength, CaretPlacement.LeadingEdgeOfCharacter)
            End If

            Dim indexOfTextLineAtDistance As Integer = GetIndexOfTextLineAtDistance(num)
            Dim textLine1 As TextLine = _textLines(indexOfTextLineAtDistance)
            Dim num2 As Double = _textLineDistances(indexOfTextLineAtDistance)
            Dim characterHitFromDistance As CharacterHit = textLine1.GetCharacterHitFromDistance(num - num2)
            Dim bufferRelativePosition As Integer = GetBufferRelativePosition(characterHitFromDistance.FirstCharacterIndex)

            If bufferRelativePosition > LineSpan.End - NewlineLength Then
                Return _textView.Caret.MoveTo(LineSpan.End - NewlineLength, CaretPlacement.LeadingEdgeOfCharacter)
            End If

            Return _textView.Caret.MoveTo(bufferRelativePosition, If((characterHitFromDistance.TrailingLength <> 0), CaretPlacement.TrailingEdgeOfCharacter, CaretPlacement.LeadingEdgeOfCharacter))
        End Function

        Public Function GetTextElementIndex(textBufferIndex As Integer) As Integer Implements ITextLine.GetTextElementIndex
            If IsDisposed Then
                Throw New ObjectDisposedException("TextLineVisual")
            End If

            If textBufferIndex < LineSpan.Start OrElse textBufferIndex > LineSpan.End - NewlineLength Then
                Throw New ArgumentOutOfRangeException("textBufferIndex")
            End If

            Dim avalonTextElementSpan As Span = GetAvalonTextElementSpan(textBufferIndex)
            Dim result = _textElementSpans.IndexOf(avalonTextElementSpan)

            If result = -1 Then
                result = (If((avalonTextElementSpan.Start <> LineSpan.Start), (GetTextElementIndex(GetAvalonTextElementSpan(avalonTextElementSpan.Start - 1).Start) + 1), 0))
                _textElementSpans.Add(avalonTextElementSpan)
            End If

            Return result
        End Function

        Public Function GetTextElementSpan(textElementIndex As Integer) As Span Implements ITextLine.GetTextElementSpan
            If IsDisposed Then
                Throw New ObjectDisposedException("TextLineVisual")
            End If

            If textElementIndex < 0 OrElse textElementIndex > GetTextElementIndex(LineSpan.End - NewlineLength) Then
                Throw New ArgumentOutOfRangeException("textElementIndex")
            End If

            Return CType(_textElementSpans(textElementIndex), Span)
        End Function

        Public Function GetPositionFromXCoordinate(x1 As Double) As Integer? Implements ITextLine.GetPositionFromXCoordinate
            If IsDisposed Then
                Throw New ObjectDisposedException("TextLineVisual")
            End If

            If Double.IsNaN(x1) Then
                Throw New ArgumentOutOfRangeException("x")
            End If

            Dim num As Double = x1 - Left
            If num < 0.0 Then Return Nothing

            If num >= (If((NewlineLength = 0), Width, (Width - 10.0))) Then
                Return Nothing
            End If

            Dim indexOfTextLineAtDistance As Integer = GetIndexOfTextLineAtDistance(num)
            Dim textLine1 As TextLine = _textLines(indexOfTextLineAtDistance)
            Dim num2 As Double = _textLineDistances(indexOfTextLineAtDistance)
            Dim bufferRelativePosition As Integer = GetBufferRelativePosition(textLine1.GetCharacterHitFromDistance(num - num2).FirstCharacterIndex)

            If bufferRelativePosition >= LineSpan.End - NewlineLength Then
                Return Nothing
            End If

            Return bufferRelativePosition
        End Function

        Public Function GetCharacterBounds(textBufferIndex As Integer) As TextBounds Implements ITextLine.GetCharacterBounds
            If IsDisposed Then
                Throw New ObjectDisposedException("TextLineVisual")
            End If

            If textBufferIndex < LineSpan.Start OrElse textBufferIndex > LineSpan.End Then
                Throw New ArgumentOutOfRangeException("textBufferIndex")
            End If

            Dim lineRelativePos = GetLineRelativePosition(textBufferIndex)
            Dim indexOfLineContaining = GetIndexOfLineContaining(lineRelativePos)
            Dim textLine1 = _textLines(indexOfLineContaining)
            Dim num As Integer = _textLineStartIndices(indexOfLineContaining)
            Dim num2 As Double
            Dim distance As Double

            If lineRelativePos > 0 AndAlso lineRelativePos = num + textLine1.Length - textLine1.NewlineLength Then
                distance = textLine1.GetDistanceFromCharacterHit(New CharacterHit(lineRelativePos - 1, 1))
                num2 = distance
            Else
                num2 = textLine1.GetDistanceFromCharacterHit(New CharacterHit(lineRelativePos, 0))
                distance = textLine1.GetDistanceFromCharacterHit(New CharacterHit(lineRelativePos, 1))
            End If

            Return New TextBounds(Left + _textLineDistances(indexOfLineContaining) + num2, Top, distance - num2, Height)
        End Function

        Public Function ContainsPosition(textBufferIndex As Integer) As Boolean Implements ITextLine.ContainsPosition
            If IsDisposed Then
                Throw New ObjectDisposedException("TextLineVisual")
            End If

            If textBufferIndex < LineSpan.Start Then Return False

            Dim [end] As Integer = LineSpan.End
            If textBufferIndex > [end] Then Return False

            If textBufferIndex = [end] AndAlso ([end] <> _textView.TextSnapshot.Length OrElse NewlineLength <> 0) Then
                Return False
            End If

            Return True
        End Function

        Public Function GetTextBounds(span1 As Span) As ReadOnlyCollection(Of TextBounds) Implements ITextLine.GetTextBounds
            If IsDisposed Then
                Throw New ObjectDisposedException("TextLineVisual")
            End If

            If span1.End > _textView.TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            Dim list1 As New List(Of TextBounds)
            Dim span2 As Span = span1

            If Not LineSpan.Contains(span1) Then
                Dim span3 = LineSpan.Overlap(span1)
                If Not span3.HasValue Then Return list1.AsReadOnly()
                span2 = span3.Value
            End If

            Dim lineRelativePosition As Integer = GetLineRelativePosition(span2.Start)
            Dim lineRelativePosition2 As Integer = GetLineRelativePosition(span2.End)
            Dim indexOfLineContaining As Integer = GetIndexOfLineContaining(lineRelativePosition)

            For i As Integer = indexOfLineContaining To _textLines.Count - 1
                Dim textLine1 As TextLine = _textLines(i)
                Dim num As Integer = _textLineStartIndices(i)
                Dim horizontalOffset As Double = _textLineDistances(i)

                Dim num2 As Integer = lineRelativePosition
                If num2 < num Then num2 = num

                Dim num3 As Integer = lineRelativePosition2
                If num3 > num + textLine1.Length Then
                    num3 = num + textLine1.Length
                End If

                list1.AddRange(GetTextBoundsOnLine(textLine1, horizontalOffset, num2, num3))
                If lineRelativePosition2 <= num + textLine1.Length Then
                    Exit For
                End If
            Next

            If span1.End >= LineSpan.End AndAlso NewlineLength > 0 Then
                list1.Add(New TextBounds(Left + Width - 10.0, Top, 10.0, Height))
            End If

            Return list1.AsReadOnly()
        End Function

        Public Function GetIndentation() As Double
            If IsDisposed Then
                Throw New ObjectDisposedException("TextLineVisual")
            End If

            Dim text As String = _textView.TextSnapshot.GetText(LineSpan)
            Dim result As Double = Width

            For i As Integer = 0 To text.Length - 1
                If Not Char.IsWhiteSpace(text(i)) Then
                    result = GetCharacterBounds(i + LineSpan.Start).Left
                    Exit For
                End If
            Next
            Return result
        End Function

        Friend Sub New(textView As ITextView, textLines As IList(Of TextLine), sourceIndex As Integer, lineSpan1 As Span, newLineLength1 As Integer, virtualCharacterPositions As IList(Of Integer), horizontalOffset As Double, verticalOffset As Double)
            _textView = textView
            _textLines = textLines
            _sourceIndex = sourceIndex
            _newLineLength = newLineLength1
            _horizontalOffset = horizontalOffset
            _verticalOffset = verticalOffset
            _span = textView.TextSnapshot.CreateTextSpan(lineSpan1, SpanTrackingMode.EdgeInclusive)
            _transform = New TranslateTransform(Left, Top)
            _textElementSpans = New ArrayList
            MyBase.Transform = _transform
            _virtualCharacterPositions = New List(Of Integer)

            For Each virtualCharacterPosition As Integer In virtualCharacterPositions
                If virtualCharacterPosition >= lineSpan1.Start Then
                    If virtualCharacterPosition > lineSpan1.End Then
                        Exit For
                    End If
                    _virtualCharacterPositions.Add(virtualCharacterPosition)
                End If
            Next

            ComputeLineProperties()
            RenderText()
        End Sub

        Friend Sub RenderText()
            Dim drawingContext1 As DrawingContext = RenderOpen()
            Dim num As Double = 0.0

            For Each textLine1 As TextLine In _textLines
                Dim y1 As Double = Baseline - textLine1.Baseline + 1.0
                textLine1.Draw(drawingContext1, New Point(num, y1), InvertAxes.None)
                num += textLine1.WidthIncludingTrailingWhitespace
            Next

            drawingContext1.Close()
        End Sub

        Friend Function GetAvalonTextElementSpan(bufferPosition As Integer) As Span
            If IsDisposed Then
                Throw New ObjectDisposedException("TextLineVisual")
            End If

            If Not ContainsPosition(bufferPosition) Then
                Throw New ArgumentOutOfRangeException("bufferPosition")
            End If

            If bufferPosition < LineSpan.Start OrElse bufferPosition > LineSpan.End Then
                Throw New ArgumentOutOfRangeException("bufferPosition")
            End If

            If bufferPosition > LineSpan.End - NewlineLength Then
                Return New Span(LineSpan.End, 0)
            End If

            If bufferPosition = LineSpan.End - NewlineLength Then
                Return New Span(LineSpan.End - NewlineLength, NewlineLength)
            End If

            Dim num As Integer = GetLineRelativePosition(bufferPosition)
            If num < 0 Then num = 0

            Dim indexOfLineContaining As Integer = GetIndexOfLineContaining(num)
            Dim textLine1 As TextLine = _textLines(indexOfLineContaining)
            Dim nextCaretCharacterHit As CharacterHit = textLine1.GetNextCaretCharacterHit(New CharacterHit(num, 0))
            Dim bufferRelativePosition As Integer = GetBufferRelativePosition(nextCaretCharacterHit.FirstCharacterIndex + nextCaretCharacterHit.TrailingLength)
            Dim bufferRelativePosition2 As Integer = GetBufferRelativePosition(textLine1.GetPreviousCaretCharacterHit(New CharacterHit(GetLineRelativePosition(bufferRelativePosition), 0)).FirstCharacterIndex)

            Return New Span(bufferRelativePosition2, bufferRelativePosition - bufferRelativePosition2)
        End Function

        Private Sub ComputeLineProperties()
            _height = 0.0
            _width = 0.0
            _extent = 0.0
            _overhangAfter = 0.0
            _baseline = 0.0

            Dim num As Double = 0.0
            Dim num2 As Integer = _sourceIndex

            _textLineDistances = New List(Of Double)
            _textLineStartIndices = New List(Of Integer)

            For Each textLine1 As TextLine In _textLines
                If textLine1.Height > _height Then
                    _height = textLine1.Height
                End If

                If textLine1.Baseline > _baseline Then
                    _baseline = textLine1.Baseline
                End If

                If textLine1.Extent > _extent Then
                    _extent = textLine1.Extent
                End If

                If textLine1.OverhangAfter > _overhangAfter Then
                    _overhangAfter = textLine1.OverhangAfter
                End If

                _width += textLine1.WidthIncludingTrailingWhitespace
                _textLineDistances.Add(num)
                num += textLine1.WidthIncludingTrailingWhitespace
                _textLineStartIndices.Add(num2)
                num2 += textLine1.Length
            Next

            _height += 2.0
            If _newLineLength > 0 Then _width += 10.0
        End Sub

        Private Function GetIndexOfTextLineAtDistance(offset As Double) As Integer
            For num As Integer = _textLineDistances.Count - 1 To 0 Step -1
                If _textLineDistances(num) <= offset Then Return num
            Next
            Return 0
        End Function

        Private Function GetIndexOfLineContaining(lineRelativePosition As Integer) As Integer
            For num As Integer = _textLineStartIndices.Count - 1 To 0 Step -1
                If _textLineStartIndices(num) <= lineRelativePosition Then
                    Return num
                End If
            Next

            Return 0
        End Function

        Private Function GetLineRelativePosition(bufferRelativePosition As Integer) As Integer
            Dim num As Integer = bufferRelativePosition - LineSpan.Start + _sourceIndex
            For Each pos In _virtualCharacterPositions
                If pos > bufferRelativePosition Then Return num
                num += 1
            Next

            Return num
        End Function

        Private Function GetBufferRelativePosition(lineRelativePosition As Integer) As Integer
            Dim num As Integer = lineRelativePosition + LineSpan.Start - _sourceIndex
            For Each virtualCharacterPosition As Integer In _virtualCharacterPositions
                If virtualCharacterPosition >= num Then Return num
                num -= 1
            Next

            Return num
        End Function

        Private Function GetTextBoundsOnLine(textLine1 As TextLine, horizontalOffset As Double, startIndex As Integer, endIndex As Integer) As List(Of TextBounds)
            Dim list1 As New List(Of TextBounds)
            If startIndex = endIndex Then
                Dim distanceFromCharacterHit As Double = textLine1.GetDistanceFromCharacterHit(New CharacterHit(startIndex, 0))
                list1.Add(New TextBounds(distanceFromCharacterHit + horizontalOffset + Left, Top, 0.0, Height))
                Return list1
            End If

            For Each textBound As System.Windows.Media.TextFormatting.TextBounds In textLine1.GetTextBounds(startIndex, endIndex - startIndex)
                list1.Add(New TextBounds(textBound.Rectangle.Left + horizontalOffset + Left, Top, textBound.Rectangle.Width, Height))
            Next

            Return list1
        End Function

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.InvariantCulture, "{0}:{1}", New Object(1) {Top, LineSpan})
        End Function

    End Class
End Namespace
