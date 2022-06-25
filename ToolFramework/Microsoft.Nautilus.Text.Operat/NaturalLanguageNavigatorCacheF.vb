Imports System.ComponentModel.Composition

Namespace Microsoft.Nautilus.Text.Operations
    Public NotInheritable Class NaturalLanguageNavigatorCacheFactory

        <Export(GetType(TextStructureNavigatorCacheFactory))>
        <ContentType("text")>
        Public Function GetTextStructureNavigator(textBuffer As ITextBuffer) As ITextStructureNavigator
            If textBuffer Is Nothing Then
                Throw New ArgumentNullException("textBuffer")
            End If

            Dim typeFromHandle As Type = GetType(NaturalLanguageNavigatorCacheFactory)
            Dim [property] As ITextStructureNavigator = Nothing

            If Not textBuffer.Properties.TryGetProperty(Of ITextStructureNavigator)(typeFromHandle, [property]) Then
                [property] = New NaturalLanguageNavigator(textBuffer)
                textBuffer.Properties.AddProperty(typeFromHandle, [property])
            End If

            Return [property]
        End Function
    End Class
End Namespace
