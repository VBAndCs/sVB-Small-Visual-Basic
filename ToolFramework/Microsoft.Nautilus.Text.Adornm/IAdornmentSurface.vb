Imports System.Collections.Generic
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.AdornmentSystem
    Public Interface IAdornmentSurface
        ReadOnly Property CanNegotiateSurfaceSpace As Boolean

        ReadOnly Property SurfacePosition As SurfacePosition

        ReadOnly Property SurfacePanel As Panel

        Sub AddAdornment(adornment As IAdornment)

        Sub RemoveAdornment(adornment As IAdornment)

        Function GetSpaceNegotiations(textLine As ITextLine) As IList(Of SpaceNegotiation)

        Function GetAdornmentsGeometry() As Geometry
    End Interface
End Namespace
