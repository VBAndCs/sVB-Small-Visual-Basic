Imports System.Collections.Generic
Imports System.ComponentModel

Namespace Microsoft.Nautilus.Core
    Public Interface IOrderable
        ReadOnly Property Name As String

        <DefaultValue(DirectCast(Nothing, Object))>
        ReadOnly Property Before As IEnumerable(Of String)

        <DefaultValue(DirectCast(Nothing, Object))>
        ReadOnly Property After As IEnumerable(Of String)
    End Interface
End Namespace
