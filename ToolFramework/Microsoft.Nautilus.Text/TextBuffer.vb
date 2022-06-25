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
            Dim normalizedTextChangeCollection1 As New NormalizedTextChangeCollection(changes)
            For Each item As TextChange In normalizedTextChangeCollection1
                builder = builder.Replace(New Span(item.Position, item.OldLength), item.NewText)
            Next
            Return normalizedTextChangeCollection1
        End Function

        Protected Overrides Function TakeSnapshot() As ITextSnapshot
            Return CreateSnapshot()
        End Function
    End Class
End Namespace
