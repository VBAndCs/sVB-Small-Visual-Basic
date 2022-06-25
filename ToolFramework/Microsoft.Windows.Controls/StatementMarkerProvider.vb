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
            _statementMarkerList.Add(marker)
            RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(marker.Span))
        End Sub

        Public Sub AddStatementMarkers(markers1 As IEnumerable(Of StatementMarker))
            Dim textSnapshot1 As ITextSnapshot = _textView.TextSnapshot
            Dim num As Integer = _textView.TextSnapshot.Length
            Dim num2 As Integer = 0

            For Each marker As StatementMarker In markers1
                Dim span1 As Span = marker.Span.GetSpan(textSnapshot1)
                If num > span1.Start Then
                    num = span1.Start
                End If

                If num2 < span1.End Then
                    num2 = span1.End
                End If

                _statementMarkerList.Add(marker)
            Next

            If num < num2 Then
                RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(New TextSpan(textSnapshot1, num, num2 - num, SpanTrackingMode.EdgeInclusive)))
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
            Dim [property] As StatementMarkerProvider = Nothing
            If Not textView.Properties.TryGetProperty(GetType(StatementMarkerProvider), [property]) Then
                [property] = New StatementMarkerProvider(textView)
                textView.Properties.AddProperty(GetType(StatementMarkerProvider), [property])
            End If

            Return [property]
        End Function

        Public Sub RemoveMarker(marker As StatementMarker)
            _statementMarkerList.Remove(marker)
            RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(marker.Span))
        End Sub
    End Class
End Namespace
