Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Text.Projection

Namespace Microsoft.Nautilus.Text.Classification
    Public NotInheritable Class ProjectionClassifierProvider

        <Import>
        Public Property ClassifierAggregatorProvider As IClassifierAggregatorProvider

        <ContentType("projection")>
        <Export(GetType(ClassifierProvider))>
        Public Function CreateClassifier(textBuffer As ITextBuffer) As IClassifier
            If TypeOf textBuffer Is IProjectionBuffer Then
                Return New ProjectionClassifier(CType(textBuffer, IProjectionBuffer), AddressOf GetAggregator)
            End If

            Throw New ArgumentException("The projection classifier can only operate on a ProjectionBuffer.")
        End Function

        Private Function GetAggregator(textBuffer As ITextBuffer) As IClassifier
            Return ClassifierAggregatorProvider.GetClassifier(textBuffer)
        End Function
    End Class
End Namespace
