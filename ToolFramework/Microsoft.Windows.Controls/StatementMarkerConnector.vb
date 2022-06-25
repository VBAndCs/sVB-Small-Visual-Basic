Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Windows.Controls
    Public NotInheritable Class StatementMarkerConnector

        <ContentType("text")>
        <Export(GetType(AdornmentProviderFactory))>
        Public Function CreateAdornmentProvider(textView As ITextView) As IAdornmentProvider
            Return StatementMarkerProvider.GetStatementMarkerProvider(textView)
        End Function

        <Export(GetType(AdornmentSurfaceFactory))>
        <AdornmentDiscriminator(GetType(StatementMarker))>
        Public Shared Function CreateAdornmentSurface(textView As IAvalonTextView) As IAdornmentSurface
            Return New StatementMarkerAdornmentSurface(textView)
        End Function
    End Class
End Namespace
