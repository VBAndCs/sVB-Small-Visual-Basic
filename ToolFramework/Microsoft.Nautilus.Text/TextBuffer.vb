Imports System.Collections.Generic
Imports Microsoft.Nautilus.Text.StringRebuilder

Namespace Microsoft.Nautilus.Text
    Friend Class TextBuffer
        Inherits BaseBuffer

        Private builder As IStringRebuilder

        Private Function CreateSnapshot() As ITextSnapshot
            Return New TextSnapshot(Me, MyBase.CurrentVersion, builder)
        End Function

        Public Sub New(contentType As String, content As IStringRebuilder)
            MyBase.New(contentType)
            builder = content
            MyBase.CurrentSnapshot = CreateSnapshot()
        End Sub

        Protected Overrides Function ApplyChanges(changes As List(Of TextChange), sourceToken As Object) As NormalizedTextChangeCollection
            Dim normalizedChanges As New NormalizedTextChangeCollection(changes)
            For Each change In normalizedChanges
                builder = builder.Replace(
                    New Span(change.Position, change.OldLength),
                    change.NewText)
            Next
            Return normalizedChanges
        End Function

        Protected Overrides Function TakeSnapshot() As ITextSnapshot
            Return CreateSnapshot()
        End Function
    End Class
End Namespace
