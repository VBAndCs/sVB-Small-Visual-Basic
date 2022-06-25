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

        Public Sub AddFindMarkers(markers As IEnumerable(Of FindMarker))
            Dim textSnapshot1 As ITextSnapshot = _textView.TextSnapshot
            Dim num As Integer = _textView.TextSnapshot.Length
            Dim num2 As Integer = 0

            For Each marker As FindMarker In markers
                Dim span1 As Span = marker.Span.GetSpan(textSnapshot1)
                If num > span1.Start Then
                    num = span1.Start
                End If

                If num2 < span1.End Then
                    num2 = span1.End
                End If

                _findMarkerList.Add(marker)
            Next

            If num < num2 Then
                RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(New TextSpan(textSnapshot1, num, num2 - num, SpanTrackingMode.EdgeInclusive)))
            End If
        End Sub

        Public Sub ClearAllMarkers()
            _findMarkerList.Clear()
            Dim changeSpan As New TextSpan(_textView.TextSnapshot, 0, _textView.TextSnapshot.Length, SpanTrackingMode.EdgeInclusive)
            RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(changeSpan))
        End Sub

        Public Function GetAdornments(span As SnapshotSpan) As IList(Of IAdornment) Implements IAdornmentProvider.GetAdornments
            Dim adornments As New List(Of IAdornment)
            For Each marker As FindMarker In _findMarkerList
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
