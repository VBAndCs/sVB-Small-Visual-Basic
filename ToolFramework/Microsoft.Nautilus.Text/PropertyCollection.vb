Imports System.Collections
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

Namespace Microsoft.Nautilus.Text
    Public Class PropertyCollection
        Implements IEnumerable(Of KeyValuePair(Of Object, Object)), IEnumerable

        Private properties As Dictionary(Of Object, Object)

        Public Sub AddProperty(key As Object, [property] As Object)
            If properties Is Nothing Then
                properties = New Dictionary(Of Object, Object)
            End If
            properties.Add(key, [property])
        End Sub

        Public Function RemoveProperty(key As Object) As Boolean
            If properties Is Nothing Then
                Return False
            End If
            Return properties.Remove(key)
        End Function

        Public Function TryGetProperty(Of TProperty)(key As Object, <Out> ByRef [property] As TProperty) As Boolean
            If properties IsNot Nothing Then
                Dim _value As Object = Nothing
                Dim flag As Boolean = properties.TryGetValue(key, _value)
                [property] = (If(flag, CType(_value, TProperty), CType(Nothing, TProperty)))
                Return flag
            End If

            [property] = CType(Nothing, TProperty)
            Return False
        End Function

        Public Iterator Function GetEnumerator() As IEnumerator(Of KeyValuePair(Of Object, Object)) Implements Collections.Generic.IEnumerable(Of Collections.Generic.KeyValuePair(Of Object, Object)).GetEnumerator
            If properties Is Nothing Then Return

            For Each [property] As KeyValuePair(Of Object, Object) In properties
                Yield [property]
            Next
        End Function

        Private Function GetEnumerator1() As IEnumerator Implements Collections.IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function
    End Class
End Namespace
