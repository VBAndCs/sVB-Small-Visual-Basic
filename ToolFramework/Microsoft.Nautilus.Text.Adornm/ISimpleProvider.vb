Imports Microsoft.Nautilus.Text.AdornmentSystem

Namespace Microsoft.Nautilus.Text.Adornments
    Public Interface ISimpleProvider
        Inherits IAdornmentProvider
        Function AddAdornment(Of T As IAdornment)(adornment As T) As T

        Function RemoveAdornment(adornment As IAdornment) As Boolean
    End Interface
End Namespace
