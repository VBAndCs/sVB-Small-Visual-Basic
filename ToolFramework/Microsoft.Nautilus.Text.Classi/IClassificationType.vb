Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text.Classification
    Public Interface IClassificationType
        ReadOnly Property Classification As String

        ReadOnly Property BaseTypes As IEnumerable(Of IClassificationType)

        Function IsOfType(type As String) As Boolean
    End Interface
End Namespace
