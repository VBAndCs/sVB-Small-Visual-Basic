Imports System.Collections.Generic
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.AdornmentSystem
    Public Interface IAdornmentSurfaceManager
        Sub AddAdornment(adornment As IAdornment)

        Sub RemoveAdornment(adornment As IAdornment)

        Function GetSpaceNegotiations(textLine As ITextLine) As IList(Of SpaceNegotiation)
    End Interface
End Namespace
