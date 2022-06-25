Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Adornments
    Public Class AvalonSimpleSurfaceService

        <Import>
        Public Property AvalonVisualFactories As IEnumerable(Of ImportInfo(Of AvalonVisualFactory, IAvalonVisualFactoryMetadata))

        <AdornmentDiscriminator(GetType(SimpleAdornment))>
        <Export(GetType(AdornmentSurfaceFactory))>
        Public Function CreateAdornmentSurface(avalonTextView As IAvalonTextView) As IAdornmentSurface
            If avalonTextView Is Nothing Then
                Throw New ArgumentNullException("avalonTextView")
            End If

            Return New AvalonSimpleSurface(avalonTextView, _avalonVisualFactories)
        End Function
    End Class
End Namespace
