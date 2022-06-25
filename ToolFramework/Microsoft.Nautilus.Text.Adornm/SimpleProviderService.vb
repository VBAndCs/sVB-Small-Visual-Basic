Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Adornments

    <Export(GetType(ISimpleProviderViewFactory))>
    <Export>
    Public NotInheritable Class SimpleProviderService
        Implements ISimpleProviderViewFactory

        Public Function GetSimpleProvider(textView As ITextView) As ISimpleProvider Implements ISimpleProviderViewFactory.GetSimpleProvider
            Return CreateSimpleAdornmentProviderViewInternal(textView)
        End Function

        <Export(GetType(AdornmentProviderFactory))>
        <ContentType("text")>
        Public Function CreateAdornmentProviderView(textView As ITextView) As IAdornmentProvider
            Return CreateSimpleAdornmentProviderViewInternal(textView)
        End Function

        Friend Function CreateSimpleAdornmentProviderViewInternal(textView As ITextView) As SimpleAdornmentProvider
            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            Dim [property] As SimpleAdornmentProvider = Nothing
            If Not textView.Properties.TryGetProperty(Of SimpleAdornmentProvider)(GetType(SimpleProviderService), [property]) Then
                [property] = New SimpleAdornmentProvider(textView.TextBuffer)
                textView.Properties.AddProperty(GetType(SimpleProviderService), [property])
            End If

            Return [property]
        End Function
    End Class
End Namespace
