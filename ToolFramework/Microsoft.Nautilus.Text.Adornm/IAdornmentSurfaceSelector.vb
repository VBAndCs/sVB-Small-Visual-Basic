Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.AdornmentSystem
    Public Interface IAdornmentSurfaceSelector
        Function CreateAdornmentSurface(textView As IAvalonTextView, adornmentType As Type) As IAdornmentSurface
    End Interface
End Namespace
