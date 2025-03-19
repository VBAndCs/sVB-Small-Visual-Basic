Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.TextFormatting
Imports Microsoft.Nautilus.Text.StringRebuilder

Namespace Microsoft.Nautilus.Text.Editor

    Friend NotInheritable Class TextLineVisual
        Inherits DrawingVisual
        Implements ITextLine, IDisposable

        Private Const _newLineWidth As Double = 10.0
        Private Const _lineHeightPadding As Double = 2.0
        Private _textView As AvalonTextView
        Private _span As ITextSpan
        Private _sourceIndex As Integer
        Private _textLines As IList(Of TextLine)
        Private _textLineDistances As List(Of Double)
        Private _textLineStartIndices As List(Of Integer)
        Private _lineLength As Integer
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

        Public ReadOnly Property LineLength As Integer Implements ITextLine.LineLength
            Get
                If IsDisposed Then
                    Throw New ObjectDisposedException("TextLineVisual")
                End If

                Return _lineLength
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
                    For Each line In _textLines
                        For Each indexedGlyphRun In line.GetIndexedGlyphRuns()
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

            If LineLength = 0 Then
                If num >= Width Then
                    Return _textView.Caret.MoveTo(LineSpan.End, CaretPlacement.LeadingEdgeOfCharacter)
                End If
            ElseIf num >= Width - 10.0 Then
                Return _textView.Caret.MoveTo(LineSpan.End - LineLength, CaretPlacement.LeadingEdgeOfCharacter)
            End If

            Dim index = GetIndexOfTextLineAtDistance(num)
            Dim textLine = _textLines(index)
            Dim distance = _textLineDistances(index)
            Dim charHit = textLine.GetCharacterHitFromDistance(num - distance)
            Dim pos = GetBufferRelativePosition(charHit.FirstCharacterIndex)

            If pos > LineSpan.End - LineLength Then
                Return _textView.Caret.MoveTo(LineSpan.End - LineLength, CaretPlacement.LeadingEdgeOfCharacter)
            End If

            Return _textView.Caret.MoveTo(
                pos,
                If(charHit.TrailingLength <> 0,
                    CaretPlacement.TrailingEdgeOfCharacter,
                    CaretPlacement.LeadingEdgeOfCharacter)
            )
        End Function

        Public Function GetTextElementIndex(textBufferIndex As Integer) As Integer Implements ITextLine.GetTextElementIndex
            If IsDisposed Then
                Throw New ObjectDisposedException("TextLineVisual")
            End If

            If textBufferIndex < LineSpan.Start OrElse textBufferIndex > LineSpan.End - LineLength Then
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

            If textElementIndex < 0 OrElse textElementIndex > GetTextElementIndex(LineSpan.End - LineLength) Then
                Throw New ArgumentOutOfRangeException("textElementIndex")
            End If

            Return CType(_textElementSpans(textElementIndex), Span)
        End Function

        Public Function GetPositionFromXCoordinate(x As Double) As Integer? Implements ITextLine.GetPositionFromXCoordinate
            If IsDisposed Then
                Throw New ObjectDisposedException("TextLineVisual")
            End If

            If Double.IsNaN(x) Then
                Throw New ArgumentOutOfRangeException("x")
            End If

            Dim margin = x - Left
            If margin < 0.0 Then Return Nothing

            If margin >= (If((LineLength = 0), Width, (Width - 10.0))) Then
                Return Nothing
            End If

            Dim index = GetIndexOfTextLineAtDistance(margin)
            Dim textLine = _textLines(index)
            Dim distance = _textLineDistances(index)
            Dim pos = GetBufferRelativePosition(textLine.GetCharacterHitFromDistance(margin - distance).FirstCharacterIndex)

            If pos >= LineSpan.End - LineLength Then
                Return Nothing
            End If

            Return pos
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

            If textBufferIndex = [end] AndAlso ([end] <> _textView.TextSnapshot.Length OrElse LineLength <> 0) Then
                Return False
            End If

            Return True
        End Function

        Public Function GetTextBounds(span As Span) As ReadOnlyCollection(Of TextBounds) Implements ITextLine.GetTextBounds
            If IsDisposed Then
                Throw New ObjectDisposedException("TextLineVisual")
            End If

            If span.End > _textView.TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            Dim textBounds As New List(Of TextBounds)
            Dim span2 As Span = span

            If Not LineSpan.Contains(span) Then
                Dim overlap = LineSpan.Overlap(span)
                If Not overlap.HasValue Then Return textBounds.AsReadOnly()
                span2 = overlap.Value
            End If

            Dim startPos = GetLineRelativePosition(span2.Start)
            Dim endPos = GetLineRelativePosition(span2.End)
            Dim lineIndex = GetIndexOfLineContaining(startPos)

            For i As Integer = lineIndex To _textLines.Count - 1
                Dim textLine = _textLines(i)
                Dim lineStart = _textLineStartIndices(i)
                Dim horizontalOffset = _textLineDistances(i)

                Dim startIndex = startPos
                If startIndex < lineStart Then startIndex = lineStart

                Dim endIndex = endPos
                If endIndex > lineStart + textLine.Length Then
                    endIndex = lineStart + textLine.Length
                End If

                textBounds.AddRange(GetTextBoundsOnLine(
                          textLine, horizontalOffset, startIndex, endIndex))
                If endPos <= lineStart + textLine.Length Then Exit For
            Next

            If span.End >= LineSpan.End AndAlso LineLength > 0 Then
                textBounds.Add(New TextBounds(Left + Width - 10.0, Top, 10.0, Height))
            End If

            Return textBounds.AsReadOnly()
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

        Dim _lineSpan As Span

        Friend Sub New(
                      textView As AvalonTextView,
                      textLines As IList(Of TextLine),
                      sourceIndex As Integer,
                      lineSpan As Span,
                      lineLength As Integer,
                      virtualCharacterPositions As IList(Of Integer),
                      horizontalOffset As Double,
                      verticalOffset As Double)

            _lineSpan = lineSpan
            _textView = textView
            _textLines = textLines
            _sourceIndex = sourceIndex
            _lineLength = lineLength
            _horizontalOffset = horizontalOffset
            _verticalOffset = verticalOffset
            _span = textView.TextSnapshot.CreateTextSpan(lineSpan, SpanTrackingMode.EdgeInclusive)
            _transform = New TranslateTransform(Left, Top)
            _textElementSpans = New ArrayList
            MyBase.Transform = _transform
            _virtualCharacterPositions = New List(Of Integer)

            For Each pos In virtualCharacterPositions
                If pos >= lineSpan.Start Then
                    If pos > lineSpan.End Then Exit For
                    _virtualCharacterPositions.Add(pos)
                End If
            Next

            ComputeLineProperties()
            RenderText()
        End Sub

        Friend Sub RenderText()
            Dim dc = RenderOpen()
            Dim x = 0.0
            For Each line In _textLines
                Dim y = Baseline - line.Baseline + 1.0 + TextLine.LineSpacing / 2
                line.Draw(dc, New Point(x, y), InvertAxes.None)
                _textView.Editor.RaiseLineRenderedEvent(
                        _textView.TextSnapshot.GetLineFromPosition(_lineSpan.Start),
                        dc, x, y - TextLine.LineSpacing / 2)
                x += line.WidthIncludingTrailingWhitespace
            Next
            dc.Close()
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

            If bufferPosition > LineSpan.End - LineLength Then
                Return New Span(LineSpan.End, 0)
            End If

            If bufferPosition = LineSpan.End - LineLength Then
                Return New Span(LineSpan.End - LineLength, LineLength)
            End If

            Dim num As Integer = GetLineRelativePosition(bufferPosition)
            If num < 0 Then num = 0

            Dim indexOfLineContaining As Integer = GetIndexOfLineContaining(num)
            Dim line = _textLines(indexOfLineContaining)
            Dim charHit = line.GetNextCaretCharacterHit(New CharacterHit(num, 0))
            Dim pos1 = GetBufferRelativePosition(charHit.FirstCharacterIndex + charHit.TrailingLength)
            Dim pos2 = GetBufferRelativePosition(line.GetPreviousCaretCharacterHit(New CharacterHit(GetLineRelativePosition(pos1), 0)).FirstCharacterIndex)
            Return New Span(pos2, pos1 - pos2)
        End Function

        Private Sub ComputeLineProperties()
            _height = 0.0
            _width = 0.0
            _extent = 0.0
            _overhangAfter = 0.0
            _baseline = 0.0

            Dim distance = 0.0
            Dim index = _sourceIndex

            _textLineDistances = New List(Of Double)
            _textLineStartIndices = New List(Of Integer)

            For Each textLine In _textLines
                If textLine.Height > _height Then
                    _height = textLine.Height
                End If

                If textLine.Baseline > _baseline Then
                    _baseline = textLine.Baseline
                End If

                If textLine.Extent > _extent Then
                    _extent = textLine.Extent
                End If

                If textLine.OverhangAfter > _overhangAfter Then
                    _overhangAfter = textLine.OverhangAfter
                End If

                _width += textLine.WidthIncludingTrailingWhitespace
                _textLineDistances.Add(distance)
                distance += textLine.WidthIncludingTrailingWhitespace
                _textLineStartIndices.Add(index)
                index += textLine.Length
            Next

            _height += 2.0
            If _lineLength > 0 Then _width += 10.0
        End Sub

        Private Function GetIndexOfTextLineAtDistance(offset As Double) As Integer
            For i = _textLineDistances.Count - 1 To 0 Step -1
                If _textLineDistances(i) <= offset Then Return i
            Next

            Return 0
        End Function

        Private Function GetIndexOfLineContaining(lineRelativePosition As Integer) As Integer
            For i = _textLineStartIndices.Count - 1 To 0 Step -1
                If _textLineStartIndices(i) <= lineRelativePosition Then
                    Return i
                End If
            Next

            Return 0
        End Function

        Private Function GetLineRelativePosition(bufferRelativePosition As Integer) As Integer
            Dim position = bufferRelativePosition - LineSpan.Start + _sourceIndex

            For Each pos In _virtualCharacterPositions
                If pos > bufferRelativePosition Then Return position
                position += 1
            Next

            Return position
        End Function

        Private Function GetBufferRelativePosition(lineRelativePosition As Integer) As Integer
            Dim position As Integer = lineRelativePosition + LineSpan.Start - _sourceIndex
            For Each pos In _virtualCharacterPositions
                If pos >= position Then Return position
                position -= 1
            Next

            Return position
        End Function

        Private Function GetTextBoundsOnLine(textLine As TextLine, horizontalOffset As Double, startIndex As Integer, endIndex As Integer) As List(Of TextBounds)
            Dim bounds As New List(Of TextBounds)
            If startIndex = endIndex Then
                Dim distance = textLine.GetDistanceFromCharacterHit(New CharacterHit(startIndex, 0))
                bounds.Add(New TextBounds(distance + horizontalOffset + Left, Top, 0.0, Height))
                Return bounds
            End If

            For Each textBounds In textLine.GetTextBounds(startIndex, endIndex - startIndex)
                bounds.Add(New TextBounds(
                           textBounds.Rectangle.Left + horizontalOffset + Left,
                           Top,
                           textBounds.Rectangle.Width,
                           Height)
               )
            Next

            Return bounds
        End Function

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.InvariantCulture, "{0}:{1}", New Object(1) {Top, LineSpan})
        End Function

    End Class
End Namespace
