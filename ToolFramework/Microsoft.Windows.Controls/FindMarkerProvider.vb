Imports System.Collections.Generic
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Windows.Controls
    Public NotInheritable Class FindMarkerProvider
        Implements IAdornmentProvider

        Private _findMarkerList As New HashSet(Of FindMarker)
        Private _textView As ITextView

        Public Event AdornmentsChanged As EventHandler(Of AdornmentsChangedEventArgs) Implements IAdornmentProvider.AdornmentsChanged

        Public Sub New(textView As ITextView)
            _textView = textView
        End Sub

        Public Sub AddFindMarker(marker As FindMarker)
            _findMarkerList.Add(marker)
            RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(marker.Span))
        End Sub

        Public Function IsHighlighted(pos As Integer) As Boolean
            Dim n = _findMarkerList.Count
            If n = 0 Then Return False

            Dim snapshot = _textView.TextSnapshot

            For i = 0 To n - 1
                Dim span = _findMarkerList(i).Span
                Dim start = span.GetStartPoint(snapshot).Position
                Dim [end] = span.GetEndPoint(snapshot).Position
                If pos >= start AndAlso pos <= [end] Then Return True
            Next
            Return False
        End Function

        Public Function GetFindMarker(moveDown As Boolean) As FindMarker
            Dim n = _findMarkerList.Count
            If n = 0 Then Return Nothing
            Dim pos = _textView.Caret.Position.TextInsertionIndex
            Dim snapshot = _textView.TextSnapshot
            Dim index = -1

            Dim i As Integer
            For i = If(moveDown, 0, n - 1) To If(moveDown, n - 1, 0) Step If(moveDown, 1, -1)
                Dim span = _findMarkerList(i).Span
                Dim start = span.GetStartPoint(snapshot).Position
                Dim [end] = span.GetEndPoint(snapshot).Position
                If (moveDown AndAlso pos < start) OrElse (Not moveDown AndAlso pos > [end]) Then
                    index = i
                    Exit For
                End If
            Next

            If i = -1 Then
                index = n - 1
            ElseIf i = n Then
                index = 0
            End If
            Return _findMarkerList(index)
        End Function

        Public Sub AddFindMarkers(markers As IEnumerable(Of FindMarker))
            Dim textSnapshot = _textView.TextSnapshot
            Dim length = _textView.TextSnapshot.Length
            Dim pos = 0

            For Each marker In markers
                Dim span = marker.Span.GetSpan(textSnapshot)
                If length > span.Start Then length = span.Start
                If pos < span.End Then pos = span.End
                _findMarkerList.Add(marker)
            Next

            If length < pos Then
                RaiseEvent AdornmentsChanged(Me,
                         New AdornmentsChangedEventArgs(
                                 New TextSpan(textSnapshot,
                                          length, pos - length,
                                          SpanTrackingMode.EdgeInclusive
                                 )
                        )
               )
            End If
        End Sub

        Public Sub ClearAllMarkers()
            _findMarkerList.Clear()
            Dim changeSpan As New TextSpan(_textView.TextSnapshot, 0, _textView.TextSnapshot.Length, SpanTrackingMode.EdgeInclusive)
            RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(changeSpan))
        End Sub

        Public Function GetAdornments(span As SnapshotSpan) As IList(Of IAdornment) Implements IAdornmentProvider.GetAdornments
            Dim adornments As New List(Of IAdornment)
            For Each marker In _findMarkerList
                If span.OverlapsWith(marker.Span.GetSpan(span.Snapshot)) Then
                    adornments.Add(marker)
                End If
            Next

            Return adornments
        End Function

        Public Shared Function GetFindMarkerProvider(textView As ITextView) As FindMarkerProvider
            Dim [property] As FindMarkerProvider = Nothing
            If Not textView.Properties.TryGetProperty(GetType(FindMarkerProvider), [property]) Then
                [property] = New FindMarkerProvider(textView)
                textView.Properties.AddProperty(GetType(FindMarkerProvider), [property])
            End If
            Return [property]
        End Function

    End Class
End Namespace
