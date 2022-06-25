Namespace Microsoft.Nautilus.Text
    Friend Class ReadOnlySpan
        Inherits TextSpan

        Public ReadOnly Property StartEdgeInsertionMode As EdgeInsertionMode

        Public ReadOnly Property EndEdgeInsertionMode As EdgeInsertionMode


        Friend Sub New(textSnapshot As ITextSnapshot, span As Span, trackingMode As SpanTrackingMode, startEdgeInsertionMode1 As EdgeInsertionMode, endEdgeInsertionMode1 As EdgeInsertionMode)
            MyBase.New(textSnapshot, span, trackingMode)
            _startEdgeInsertionMode = startEdgeInsertionMode1
            _endEdgeInsertionMode = endEdgeInsertionMode1
        End Sub

        Public Function IsReplaceAllowed(span As Span) As Boolean
            Dim span2 As Span = GetSpan(MyBase.TextBuffer.CurrentSnapshot)
            If span2.Length = 0 Then Return True

            If span2.OverlapsWith(span) Then Return False

            If _StartEdgeInsertionMode = EdgeInsertionMode.Deny AndAlso span.Start = span2.Start Then
                Return False
            End If

            If _EndEdgeInsertionMode = EdgeInsertionMode.Deny AndAlso span.End = span2.End Then
                Return False
            End If

            Return True
        End Function

        Public Function IsInsertAllowed(position As Integer) As Boolean
            Dim span As Span = GetSpan(MyBase.TextBuffer.CurrentSnapshot)
            If span.Start < position AndAlso span.End > position Then
                Return False
            End If

            If _StartEdgeInsertionMode = EdgeInsertionMode.Deny AndAlso position = span.Start Then
                Return False
            End If

            If _EndEdgeInsertionMode = EdgeInsertionMode.Deny AndAlso position = span.End Then
                Return False
            End If

            Return True
        End Function

        Public Function Intersects(span As ReadOnlySpan) As Boolean
            Dim currentSnapshot1 As ITextSnapshot = MyBase.TextBuffer.CurrentSnapshot
            Dim span2 As Span = GetSpan(currentSnapshot1)
            Dim span3 As Span = span.GetSpan(currentSnapshot1)
            If span2.OverlapsWith(span3) Then Return True

            If (_StartEdgeInsertionMode = EdgeInsertionMode.Deny OrElse span._EndEdgeInsertionMode = EdgeInsertionMode.Deny) AndAlso span2.Start = span3.End Then
                Return True
            End If

            If (_EndEdgeInsertionMode = EdgeInsertionMode.Deny OrElse span._StartEdgeInsertionMode = EdgeInsertionMode.Deny) AndAlso span2.End = span3.Start Then
                Return True
            End If

            Return False
        End Function
    End Class
End Namespace
