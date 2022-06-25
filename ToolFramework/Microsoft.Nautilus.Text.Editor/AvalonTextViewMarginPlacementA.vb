Imports Microsoft.Nautilus.Core

Namespace Microsoft.Nautilus.Text.Editor
    Public NotInheritable Class AvalonTextViewMarginPlacementAttribute
        Inherits BaseMetadataProviderAttribute

        Public ReadOnly Property MarginPlacement As Integer

        Public Sub New(placement As MarginPlacement)
            _MarginPlacement = CInt(CLng(Fix(placement)) Mod Integer.MaxValue)
        End Sub
    End Class
End Namespace
