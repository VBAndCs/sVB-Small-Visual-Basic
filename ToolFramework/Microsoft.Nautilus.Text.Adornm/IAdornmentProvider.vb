Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text.AdornmentSystem
    Public Interface IAdornmentProvider
        Event AdornmentsChanged As EventHandler(Of AdornmentsChangedEventArgs)

        Function GetAdornments(span As SnapshotSpan) As IList(Of IAdornment)
    End Interface
End Namespace
