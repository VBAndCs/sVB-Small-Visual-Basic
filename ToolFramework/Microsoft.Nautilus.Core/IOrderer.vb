Imports System.Collections.Generic
Imports System.ComponentModel.Composition

Namespace Microsoft.Nautilus.Core
    Public Interface IOrderer
        Function Order(Of TValue, TMetadata As IOrderable)(items As IEnumerable(Of ImportInfo(Of TValue, TMetadata))) As IEnumerable(Of ImportInfo(Of TValue, TMetadata))

        Function Order(Of TOrderable As IOrderable)(items As IEnumerable(Of TOrderable)) As IEnumerable(Of TOrderable)

        Function Order(Of T)(items As IEnumerable(Of T), converter1 As Converter(Of T, IOrderable)) As IEnumerable(Of T)
    End Interface
End Namespace
