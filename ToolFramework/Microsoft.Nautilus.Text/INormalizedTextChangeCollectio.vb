Imports System.Collections
Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text
    Public Interface INormalizedTextChangeCollection
        Inherits IList(Of ITextChange), ICollection(Of ITextChange), IEnumerable(Of ITextChange), IEnumerable
    End Interface
End Namespace
