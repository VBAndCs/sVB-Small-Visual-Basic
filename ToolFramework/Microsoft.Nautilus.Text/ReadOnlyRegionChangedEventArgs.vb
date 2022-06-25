
Namespace Microsoft.Nautilus.Text
    Public Class ReadOnlyRegionChangedEventArgs
        Inherits EventArgs

        Public ReadOnly Property ReadOnlyRegion As IReadOnlyRegionHandle

        Public ReadOnly Property Change As ReadOnlyRegionChange

        Public Sub New(readOnlyRegion As IReadOnlyRegionHandle, change As ReadOnlyRegionChange)
            _ReadOnlyRegion = readOnlyRegion
            _Change = change
        End Sub
    End Class

End Namespace
