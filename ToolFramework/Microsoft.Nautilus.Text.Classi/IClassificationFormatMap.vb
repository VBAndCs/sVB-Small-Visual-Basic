Imports System.Collections.ObjectModel
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Classification
    Public Interface IClassificationFormatMap
        ReadOnly Property CurrentPriorityOrder As ReadOnlyCollection(Of IClassificationType)

        Function GetTextProperties(classificationType As IClassificationType) As TextFormattingRunProperties

        Sub SetTextProperties(classificationType As IClassificationType, properties As TextFormattingRunProperties)

        Sub SwapPriorities(firstType As IClassificationType, secondType As IClassificationType)
    End Interface
End Namespace
