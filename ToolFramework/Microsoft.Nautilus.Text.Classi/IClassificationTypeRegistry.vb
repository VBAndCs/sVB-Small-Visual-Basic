Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text.Classification
    Public Interface IClassificationTypeRegistry
        Function GetClassificationType(type As String) As IClassificationType

        Function CreateTransientClassificationType(baseTypes As IEnumerable(Of IClassificationType)) As IClassificationType

        Function CreateTransientClassificationType(ParamArray baseTypes As IClassificationType()) As IClassificationType
    End Interface
End Namespace
