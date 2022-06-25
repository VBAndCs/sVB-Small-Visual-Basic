Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.AdornmentSystem
    Public Interface IAdornmentSurfaceHost
        ReadOnly Property TextView As IAvalonTextView

        Sub AddAdornmentSurface(adornmentSurface As IAdornmentSurface)
    End Interface
End Namespace
