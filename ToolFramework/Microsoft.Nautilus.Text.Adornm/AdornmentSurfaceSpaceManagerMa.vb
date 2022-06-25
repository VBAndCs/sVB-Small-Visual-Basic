Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.AdornmentSystem
    <Export(GetType(IAdornmentSurfaceSpaceManagerMap))>
    Public NotInheritable Class AdornmentSurfaceSpaceManagerMapService
        Implements IAdornmentSurfaceSpaceManagerMap

        Public Function CreateAdornmentSurfaceSpaceManager(avalonTextView As IAvalonTextView) As IAdornmentSurfaceSpaceManager Implements IAdornmentSurfaceSpaceManagerMap.CreateAdornmentSurfaceSpaceManager
            If avalonTextView Is Nothing Then
                Throw New ArgumentNullException("avalonTextView")
            End If

            Dim [property] As IAdornmentSurfaceSpaceManager = Nothing
            If Not avalonTextView.Properties.TryGetProperty(GetType(AdornmentSurfaceManagerService), [property]) Then
                Throw New ArgumentException("The specified avalonTextView doesn't have a space manager.")
            End If

            Return [property]
        End Function
    End Class
End Namespace
