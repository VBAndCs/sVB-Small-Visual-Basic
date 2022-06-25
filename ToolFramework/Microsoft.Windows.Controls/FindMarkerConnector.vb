Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Windows.Controls

    Public NotInheritable Class FindMarkerConnector
        <ContentType("text")>
        <Export(GetType(AdornmentProviderFactory))>
        Public Function CreateAdornmentProvider(textView As ITextView) As IAdornmentProvider
            Return FindMarkerProvider.GetFindMarkerProvider(textView)
        End Function

        <AdornmentDiscriminator(GetType(FindMarker))>
        <Export(GetType(AdornmentSurfaceFactory))>
        Public Shared Function CreateAdornmentSurface(textView As IAvalonTextView) As IAdornmentSurface
            Return New FindMarkerAdornmentSurface(textView)
        End Function
    End Class
End Namespace
