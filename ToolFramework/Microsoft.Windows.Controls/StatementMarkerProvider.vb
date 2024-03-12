Imports System.Collections.Generic
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Windows.Controls
    Public NotInheritable Class StatementMarkerProvider
        Implements IAdornmentProvider

        Private _statementMarkerList As New HashSet(Of StatementMarker)
        Private _textView As ITextView

        Public ReadOnly Property Markers As HashSet(Of StatementMarker)
            Get
                Return _statementMarkerList
            End Get
        End Property

        Public Event AdornmentsChanged As EventHandler(Of AdornmentsChangedEventArgs) Implements IAdornmentProvider.AdornmentsChanged

        Public Sub New(textView As ITextView)
            _textView = textView
        End Sub

        Public Sub AddStatementMarker(marker As StatementMarker)
            marker.RefreshMarker = AddressOf RefreshMarker
            _statementMarkerList.Add(marker)
            RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(marker.Span))
        End Sub

        Private Sub RefreshMarker(marker As StatementMarker)
            RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(marker.Span))
        End Sub

        Public Sub AddStatementMarkers(markers As IEnumerable(Of StatementMarker))
            Dim snapshot = _textView.TextSnapshot
            Dim spanStart = _textView.TextSnapshot.Length
            Dim spanEnd = 0

            For Each marker As StatementMarker In markers
                Dim span = marker.Span.GetSpan(snapshot)
                If spanStart > span.Start Then
                    spanStart = span.Start
                End If

                If spanEnd < span.End Then
                    spanEnd = span.End
                End If

                _statementMarkerList.Add(marker)
            Next

            If spanStart < spanEnd Then
                RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(New TextSpan(snapshot, spanStart, spanEnd - spanStart, SpanTrackingMode.EdgeInclusive)))
            End If
        End Sub

        Public Sub ClearAllMarkers()
            _statementMarkerList.Clear()
            Dim changeSpan As New TextSpan(_textView.TextSnapshot, 0, _textView.TextSnapshot.Length, SpanTrackingMode.EdgeInclusive)
            RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(changeSpan))
        End Sub

        Public Function GetAdornments(span1 As SnapshotSpan) As IList(Of IAdornment) Implements IAdornmentProvider.GetAdornments
            Dim list1 As New List(Of IAdornment)
            For Each statementMarker1 As StatementMarker In _statementMarkerList
                If span1.OverlapsWith(statementMarker1.Span.GetSpan(span1.Snapshot)) Then
                    list1.Add(statementMarker1)
                End If
            Next

            Return list1
        End Function

        Public Shared Function GetStatementMarkerProvider(textView As ITextView) As StatementMarkerProvider
            Dim provider As StatementMarkerProvider = Nothing
            If Not textView.Properties.TryGetProperty(GetType(StatementMarkerProvider), provider) Then
                provider = New StatementMarkerProvider(textView)
                textView.Properties.AddProperty(GetType(StatementMarkerProvider), provider)
            End If

            Return provider
        End Function

        Public Sub RemoveMarker(marker As StatementMarker)
            If marker Is Nothing Then Return
            _statementMarkerList.Remove(marker)
            RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(marker.Span))
        End Sub

        Public Function GetMarker(lineNumber As Integer) As StatementMarker
            Dim snapshot = _textView.TextSnapshot
            For Each marker In _statementMarkerList
                If marker.LineNumber = lineNumber Then
                    Return marker
                End If
            Next
            Return Nothing
        End Function
    End Class
End Namespace
