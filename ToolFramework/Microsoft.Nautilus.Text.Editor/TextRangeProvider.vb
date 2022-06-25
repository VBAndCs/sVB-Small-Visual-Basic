Imports System.Collections.Generic
Imports System.Globalization
Imports System.Windows.Automation.Provider
Imports System.Windows.Automation.Text

Namespace Microsoft.Nautilus.Text.Editor.Automation.Implementation
    Friend Class TextRangeProvider
        Implements ITextRangeProvider

        Friend ReadOnly Property TextView As IAvalonTextView

        Friend ReadOnly Property StartPoint As ITextPoint

        Friend ReadOnly Property EndPoint As ITextPoint

        Public Sub New(textView As IAvalonTextView, startPoint As ITextPoint, endPoint As ITextPoint)
            _TextView = textView
            _StartPoint = startPoint
            _EndPoint = endPoint
        End Sub

        Public Sub New(textView As IAvalonTextView, span As Span)
            _TextView = textView
            _StartPoint = _TextView.TextSnapshot.CreateTextPoint(span.Start, TrackingMode.Positive)
            _EndPoint = _TextView.TextSnapshot.CreateTextPoint(span.End, TrackingMode.Positive)
        End Sub

        Private Function GetChildren() As IRawElementProviderSimple() Implements ITextRangeProvider.GetChildren
            Return Nothing
        End Function

        Private Function Clone() As ITextRangeProvider Implements ITextRangeProvider.Clone
            Return New TextRangeProvider(_textView, _startPoint, _endPoint)
        End Function

        Private Function Compare(range As ITextRangeProvider) As Boolean Implements ITextRangeProvider.Compare
            If range Is Nothing Then
                Throw New ArgumentNullException("range")
            End If

            If TypeOf range IsNot TextRangeProvider Then
                Throw New ArgumentException("Supplied range is not valid.")
            End If

            Dim textRange = CType(range, TextRangeProvider)

            If textRange.TextView Is _textView AndAlso textRange.StartPoint Is _startPoint AndAlso textRange.EndPoint Is _endPoint Then
                Return True
            End If

            Return False
        End Function

        Private Function CompareEndpoints(endpoint2 As TextPatternRangeEndpoint, targetRange As ITextRangeProvider, targetEndpoint As TextPatternRangeEndpoint) As Integer Implements ITextRangeProvider.CompareEndpoints
            If targetRange Is Nothing Then
                Throw New ArgumentNullException("targetRange")
            End If

            If TypeOf targetRange IsNot TextRangeProvider Then
                Throw New ArgumentException("Supplied target range is not valid.")
            End If

            Dim textRange = CType(targetRange, TextRangeProvider)
            Dim point = If(endpoint2 = TextPatternRangeEndpoint.Start, _startPoint, _endPoint)
            Dim position As Integer = point.GetPosition(_textView.TextSnapshot)
            point = If((targetEndpoint = TextPatternRangeEndpoint.Start), textRange.StartPoint, textRange.EndPoint)
            Dim position2 As Integer = point.GetPosition(_textView.TextSnapshot)

            Return position.CompareTo(position2)
        End Function

        Private Sub ExpandToEnclosingUnit(unit As TextUnit) Implements ITextRangeProvider.ExpandToEnclosingUnit
            Throw New NotImplementedException
        End Sub

        Private Function FindAttribute(attribute As Integer, val As Object, backward As Boolean) As ITextRangeProvider Implements ITextRangeProvider.FindAttribute
            Throw New NotImplementedException
        End Function

        Private Function FindText(text As String, backward As Boolean, ignoreCase As Boolean) As ITextRangeProvider Implements ITextRangeProvider.FindText
            If text Is Nothing Then
                Throw New ArgumentNullException("text")
            End If

            Dim text2 As String = _textView.TextSnapshot.GetText(Span.FromBounds(_startPoint.GetPosition(_textView.TextSnapshot), _endPoint.GetPosition(_textView.TextSnapshot)))
            Dim num As Integer = -1

            If ignoreCase Then
                text2 = text2.ToUpper(CultureInfo.InvariantCulture)
                text = text.ToUpper(CultureInfo.InvariantCulture)
            End If

            num = (If((Not backward), text2.IndexOf(text, StringComparison.Ordinal), text2.LastIndexOf(text, StringComparison.Ordinal)))
            If num = -1 Then
                Return Nothing
            End If

            Return New TextRangeProvider(_textView, _textView.TextSnapshot.CreateTextPoint(_startPoint.GetPosition(_textView.TextSnapshot) + num, TrackingMode.Positive), _textView.TextSnapshot.CreateTextPoint(_startPoint.GetPosition(_textView.TextSnapshot) + num + text.Length, TrackingMode.Positive))
        End Function

        Private Function GetAttributeValue(attribute As Integer) As Object Implements ITextRangeProvider.GetAttributeValue
            Throw New NotImplementedException
        End Function

        Private Function GetBoundingRectangles() As Double() Implements ITextRangeProvider.GetBoundingRectangles
            Dim textLineVisuals As List(Of TextLineVisual) = GetTextLineVisuals()
            Dim array As Double() = New Double(4 * textLineVisuals.Count - 1) {}
            For i As Integer = 0 To 4 * textLineVisuals.Count - 1 Step 4
                Dim textLineVisual1 As TextLineVisual = textLineVisuals(i \ 4)
                array(i) = textLineVisual1.Left
                array(i + 1) = textLineVisual1.Top
                array(i + 2) = textLineVisual1.Width
                array(i + 3) = textLineVisual1.Height
            Next
            Return array
        End Function

        Private Function GetEnclosingElement() As IRawElementProviderSimple Implements ITextRangeProvider.GetEnclosingElement
            Return Nothing
        End Function

        Private Function GetText(maxLength As Integer) As String Implements ITextRangeProvider.GetText
            If maxLength < -1 Then
                Throw New ArgumentOutOfRangeException("maxLength")
            End If

            If maxLength = -1 Then
                Return _textView.TextSnapshot.GetText(Span.FromBounds(_startPoint.GetPosition(_textView.TextSnapshot), _endPoint.GetPosition(_textView.TextSnapshot)))
            End If

            Return _textView.TextSnapshot.GetText(Span.FromBounds(_startPoint.GetPosition(_textView.TextSnapshot), Math.Min(_textView.TextSnapshot.Length, _startPoint.GetPosition(_textView.TextSnapshot) + maxLength)))
        End Function

        Private Function Move(unit As TextUnit, count1 As Integer) As Integer Implements ITextRangeProvider.Move
            If unit = TextUnit.Character Then
                If _startPoint.GetPosition(_textView.TextSnapshot) + count1 < 0 Then
                    count1 = -_startPoint.GetPosition(_textView.TextSnapshot)
                End If

                If _endPoint.GetPosition(_textView.TextSnapshot) + count1 > _textView.TextSnapshot.Length Then
                    count1 = _textView.TextSnapshot.Length - _endPoint.GetPosition(_textView.TextSnapshot)
                End If

                _startPoint = _textView.TextSnapshot.CreateTextPoint(_startPoint.GetPosition(_textView.TextSnapshot) + count1, TrackingMode.Positive)
                _endPoint = _textView.TextSnapshot.CreateTextPoint(_endPoint.GetPosition(_textView.TextSnapshot) + count1, TrackingMode.Positive)
                Return count1
            End If

            Throw New NotImplementedException
        End Function

        Private Sub MoveEndpointByRange(endpoint2 As TextPatternRangeEndpoint, targetRange As ITextRangeProvider, targetEndpoint As TextPatternRangeEndpoint) Implements ITextRangeProvider.MoveEndpointByRange
            If targetRange Is Nothing Then
                Throw New ArgumentNullException("targetRange")
            End If

            If TypeOf targetRange IsNot TextRangeProvider Then
                Throw New ArgumentException("Supplied target range is not valid.")
            End If

            Dim textRange2 = CType(targetRange, TextRangeProvider)

            Dim position As Integer = (If((targetEndpoint = TextPatternRangeEndpoint.Start), textRange2.StartPoint.GetPosition(_textView.TextSnapshot), textRange2.EndPoint.GetPosition(_textView.TextSnapshot)))
            If endpoint2 = TextPatternRangeEndpoint.Start Then
                _startPoint = _textView.TextSnapshot.CreateTextPoint(position, TrackingMode.Positive)
                If _startPoint.GetPosition(_textView.TextSnapshot) > _endPoint.GetPosition(_textView.TextSnapshot) Then
                    _endPoint = _textView.TextSnapshot.CreateTextPoint(_startPoint.GetPosition(_textView.TextSnapshot), TrackingMode.Positive)
                End If

            Else
                _EndPoint = _textView.TextSnapshot.CreateTextPoint(position, TrackingMode.Positive)
                If _endPoint.GetPosition(_textView.TextSnapshot) < _startPoint.GetPosition(_textView.TextSnapshot) Then
                    _startPoint = _textView.TextSnapshot.CreateTextPoint(_endPoint.GetPosition(_textView.TextSnapshot), TrackingMode.Positive)
                End If
            End If

        End Sub

        Private Function MoveEndpointByUnit(endpoint2 As TextPatternRangeEndpoint, unit As TextUnit, count1 As Integer) As Integer Implements ITextRangeProvider.MoveEndpointByUnit
            If unit = TextUnit.Character Then
                If endpoint2 = TextPatternRangeEndpoint.Start Then
                    If _StartPoint.GetPosition(_TextView.TextSnapshot) + count1 < 0 Then
                        count1 = -_StartPoint.GetPosition(_TextView.TextSnapshot)
                    ElseIf _StartPoint.GetPosition(_TextView.TextSnapshot) + count1 > _TextView.TextSnapshot.Length Then
                        count1 = _TextView.TextSnapshot.Length - _StartPoint.GetPosition(_TextView.TextSnapshot)
                    End If

                    _StartPoint = _TextView.TextSnapshot.CreateTextPoint(_StartPoint.GetPosition(_TextView.TextSnapshot) + count1, TrackingMode.Positive)
                    If _StartPoint.GetPosition(_TextView.TextSnapshot) > _EndPoint.GetPosition(_TextView.TextSnapshot) Then
                        _EndPoint = _TextView.TextSnapshot.CreateTextPoint(_StartPoint.GetPosition(_TextView.TextSnapshot), TrackingMode.Positive)
                    End If

                Else
                    If _EndPoint.GetPosition(_TextView.TextSnapshot) + count1 < 0 Then
                        count1 = -_EndPoint.GetPosition(_TextView.TextSnapshot)
                    ElseIf _EndPoint.GetPosition(_TextView.TextSnapshot) + count1 > _TextView.TextSnapshot.Length Then
                        count1 = _TextView.TextSnapshot.Length - _EndPoint.GetPosition(_TextView.TextSnapshot)
                    End If

                    _EndPoint = _TextView.TextSnapshot.CreateTextPoint(_EndPoint.GetPosition(_TextView.TextSnapshot) + count1, TrackingMode.Positive)
                    If _EndPoint.GetPosition(_TextView.TextSnapshot) < _StartPoint.GetPosition(_TextView.TextSnapshot) Then
                        _StartPoint = _TextView.TextSnapshot.CreateTextPoint(_EndPoint.GetPosition(_TextView.TextSnapshot), TrackingMode.Positive)
                    End If

                End If

                Return count1
            End If

            Throw New NotImplementedException
        End Function

        Private Sub ScrollIntoView(alignToTop As Boolean) Implements ITextRangeProvider.ScrollIntoView
            If alignToTop Then
                _textView.DisplayTextLineContainingCharacter(_startPoint.GetPosition(_textView.TextSnapshot), 0.0, ViewRelativePosition.Top)
            Else
                _textView.DisplayTextLineContainingCharacter(_endPoint.GetPosition(_textView.TextSnapshot), 0.0, ViewRelativePosition.Bottom)
            End If
        End Sub

        Private Sub [Select]() Implements ITextRangeProvider.Select
            _textView.Selection.ActiveSnapshotSpan = New SnapshotSpan(_textView.TextSnapshot, Span.FromBounds(_startPoint.GetPosition(_textView.TextSnapshot), _endPoint.GetPosition(_textView.TextSnapshot)))
        End Sub

        Public Sub AddToSelection() Implements ITextRangeProvider.AddToSelection
        End Sub

        Public Sub RemoveFromSelection() Implements ITextRangeProvider.RemoveFromSelection
        End Sub

        Private Function GetTextLineVisuals() As List(Of TextLineVisual)
            Dim list1 As New List(Of TextLineVisual)
            Dim formattedTextLines1 As IList(Of ITextLine) = _textView.FormattedTextLines

            For Each item2 As ITextLine In formattedTextLines1
                Dim start1 As Integer = item2.LineSpan.Start
                Dim [end] As Integer = item2.LineSpan.End

                If (start1 >= _StartPoint.GetPosition(_TextView.TextSnapshot) AndAlso start1 <= _EndPoint.GetPosition(_TextView.TextSnapshot)) OrElse ([end] >= _StartPoint.GetPosition(_TextView.TextSnapshot) AndAlso [end] <= _EndPoint.GetPosition(_TextView.TextSnapshot)) Then
                    Dim item As TextLineVisual = TryCast(item2, TextLineVisual)
                    If item2 IsNot Nothing Then
                        list1.Add(item)
                    End If
                End If
            Next

            Return list1
        End Function
    End Class
End Namespace
