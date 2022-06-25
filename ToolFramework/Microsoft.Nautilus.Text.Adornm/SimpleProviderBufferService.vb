Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Adornments
    <Export(GetType(ISimpleProviderBufferFactory))>
    Public NotInheritable Class SimpleProviderBufferService
        Implements ISimpleProviderBufferFactory

        Public Function GetSimpleProvider(textView As ITextView) As ISimpleProvider Implements ISimpleProviderBufferFactory.GetSimpleProvider
            Return CreateSimpleAdornmentProviderBufferInternal(textView)
        End Function

        <Export(GetType(AdornmentProviderFactory))>
        <ContentType("text")>
        Public Function CreateAdornmentProviderBuffer(textView As ITextView) As IAdornmentProvider
            Return CreateSimpleAdornmentProviderBufferInternal(textView)
        End Function

        Friend Function CreateSimpleAdornmentProviderBufferInternal(textView As ITextView) As SimpleAdornmentProvider
            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            Dim textBuffer1 As ITextBuffer = textView.TextBuffer
            Dim [property] As SimpleAdornmentProvider = Nothing
            If Not textBuffer1.Properties.TryGetProperty(Of SimpleAdornmentProvider)(GetType(SimpleProviderService), [property]) Then
                [property] = New SimpleAdornmentProvider(textView.TextBuffer)
                textBuffer1.Properties.AddProperty(GetType(SimpleProviderService), [property])
            End If
            Return [property]
        End Function
    End Class
End Namespace
