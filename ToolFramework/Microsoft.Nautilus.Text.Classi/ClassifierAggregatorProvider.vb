Imports System.Collections.Generic
Imports System.ComponentModel.Composition

Namespace Microsoft.Nautilus.Text.Classification
    <Export(GetType(IClassifierAggregatorProvider))>
    Public NotInheritable Class ClassifierAggregatorProvider
        Implements IClassifierAggregatorProvider

        <Import(GetType(ClassifierProvider))>
        Public Property ClassifierProviders As IEnumerable(Of ImportInfo(Of Func(Of ITextBuffer, IClassifier), IContentTypeMetadata))

        <Import>
        Public Property ClassificationTypeRegistry As IClassificationTypeRegistry


        Public Function GetClassifier(textBuffer As ITextBuffer) As IClassifier Implements IClassifierAggregatorProvider.GetClassifier
            If textBuffer Is Nothing Then
                Throw New ArgumentNullException("textBuffer")
            End If

            Dim aggregatorType As Type = GetType(ClassifierAggregator)
            Dim aggregator As WeakReference = Nothing

            If Not textBuffer.Properties.TryGetProperty(aggregatorType, aggregator) Then
                Return CreateAndAddAggregator(textBuffer)
            End If

            Dim target = aggregator.Target
            If target Is Nothing Then
                textBuffer.Properties.RemoveProperty(aggregatorType)
                Return CreateAndAddAggregator(textBuffer)
            End If

            Return CType(target, ClassifierAggregator)
        End Function

        Private Function CreateAndAddAggregator(textBuffer As ITextBuffer) As ClassifierAggregator
            Dim classifierAggregator1 As New ClassifierAggregator(textBuffer, ClassifierProviders, ClassificationTypeRegistry)
            Dim [property] As New WeakReference(classifierAggregator1)
            textBuffer.Properties.AddProperty(GetType(ClassifierAggregator), [property])
            Return classifierAggregator1
        End Function
    End Class
End Namespace
