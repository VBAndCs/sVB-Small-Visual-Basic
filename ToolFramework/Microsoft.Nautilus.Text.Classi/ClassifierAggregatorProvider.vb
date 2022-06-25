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

            Dim classifierAggregator1 As ClassifierAggregator = Nothing
            Dim typeFromHandle As Type = GetType(ClassifierAggregator)
            Dim [property] As WeakReference = Nothing

            If Not textBuffer.Properties.TryGetProperty(typeFromHandle, [property]) Then
                Return CreateAndAddAggregator(textBuffer)
            End If

            Dim target1 As Object = [property].Target
            If target1 Is Nothing Then
                textBuffer.Properties.RemoveProperty(typeFromHandle)
                Return CreateAndAddAggregator(textBuffer)
            End If

            Return CType(target1, ClassifierAggregator)
        End Function

        Private Function CreateAndAddAggregator(textBuffer As ITextBuffer) As ClassifierAggregator
            Dim classifierAggregator1 As New ClassifierAggregator(textBuffer, ClassifierProviders, ClassificationTypeRegistry)
            Dim [property] As New WeakReference(classifierAggregator1)
            textBuffer.Properties.AddProperty(GetType(ClassifierAggregator), [property])
            Return classifierAggregator1
        End Function
    End Class
End Namespace
