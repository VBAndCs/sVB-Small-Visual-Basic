
Namespace Microsoft.Nautilus.Text.Classification
    Public Class ClassificationSpan
        Implements ITextSpan

        Private textSpan As ITextSpan
        Private classification As IClassificationType

        Public ReadOnly Property ClassificationType As IClassificationType
            Get
                Return classification
            End Get
        End Property

        Public ReadOnly Property TextBuffer As ITextBuffer Implements ITextSpan.TextBuffer
            Get
                Return textSpan.TextBuffer
            End Get
        End Property

        Public ReadOnly Property TrackingMode As SpanTrackingMode Implements ITextSpan.TrackingMode
            Get
                Return textSpan.TrackingMode
            End Get
        End Property

        Public Sub New(textSpan As ITextSpan, classification As IClassificationType)
            If textSpan Is Nothing Then
                Throw New ArgumentNullException("textSpan")
            End If
            If classification Is Nothing Then
                Throw New ArgumentNullException("classification")
            End If
            Me.textSpan = textSpan
            Me.classification = classification
        End Sub

        Public Function GetSpan(snapshot As ITextSnapshot) As SnapshotSpan Implements ITextSpan.GetSpan
            Return textSpan.GetSpan(snapshot)
        End Function

        Public Function GetEndPoint(snapshot As ITextSnapshot) As SnapshotPoint Implements ITextSpan.GetEndPoint
            Return textSpan.GetEndPoint(snapshot)
        End Function

        Public Function GetStartPoint(snapshot As ITextSnapshot) As SnapshotPoint Implements ITextSpan.GetStartPoint
            Return textSpan.GetStartPoint(snapshot)
        End Function

        Public Function GetText(snapshot As ITextSnapshot) As String Implements ITextSpan.GetText
            Return textSpan.GetText(snapshot)
        End Function
    End Class
End Namespace
