Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.AdornmentSystem
    <Export(GetType(IAdornmentAggregatorCache))>
    Public NotInheritable Class AggregatorService
        Implements IAdornmentAggregatorCache

        <Import(GetType(AdornmentProviderFactory))>
        Public Property AdornmentProviderFactories As IEnumerable(Of ImportInfo(Of Func(Of ITextView, IAdornmentProvider), IContentTypeMetadata))

        Private Function GetAdornmentProvider(textView As ITextView) As IAdornmentProvider Implements IAdornmentAggregatorCache.GetAdornmentProvider
            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            Dim typeFromHandle As Type = GetType(AggregatorService)
            Dim [property] As IAdornmentProvider = Nothing
            If Not textView.Properties.TryGetProperty(Of IAdornmentProvider)(typeFromHandle, [property]) Then
                [property] = New AdornmentAggregator(textView, _adornmentProviderFactories)
                textView.Properties.AddProperty(typeFromHandle, [property])
            End If
            Return [property]
        End Function
    End Class
End Namespace
