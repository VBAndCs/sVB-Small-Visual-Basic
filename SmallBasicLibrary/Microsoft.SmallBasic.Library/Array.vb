' To configure or remove Option's included in result, go to Options/Advanced Options...
Option Compare Text
Option Explicit On
Option Infer On
Option Strict On
Imports System.Collections.Generic

Namespace Microsoft.SmallBasic.Library
    ''' <summary>
    ''' This object provides a way of storing more than one value for a given name. These values can be accessed by another index.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Array
        Private Shared _arrayMap As New Dictionary(Of Primitive, Dictionary(Of Primitive, Primitive))
        ''' <summary>
        ''' Gets whether or not the array contains the specified index.  This is very useful when deciding if the array's index was initialized by some value or not.
        ''' </summary>
        ''' <param name="array">
        ''' The array to check.
        ''' </param>
        ''' <param name="index">
        ''' The index to check.
        ''' </param>
        ''' <returns>
        ''' "True" or "False" depending on if the index was present in the specified.
        ''' array.
        ''' </returns>
        Public Shared Function ContainsIndex(array1 As Primitive, index As Primitive) As Primitive
            Return array1.ContainsKey(index)
        End Function
        ''' <summary>
        ''' Gets whether or not the array contains the specified value.  This is very useful when deciding if the array's value was stored in some index.
        ''' </summary>
        ''' <param name="array">
        ''' The array to check.
        ''' </param>
        ''' <param name="value">
        ''' The value to check.
        ''' </param>
        ''' <returns>
        ''' "True" or "False" depending on if the value was present in the specified.
        ''' array.
        ''' </returns>
        Public Shared Function ContainsValue(array1 As Primitive, value As Primitive) As Primitive
            Return array1.ContainsValue(value)
        End Function
        ''' <summary>
        ''' Gets all the indices for the array, as another array.
        ''' </summary>
        ''' <param name="array">
        ''' The array whose indices are requested.
        ''' </param>
        ''' <returns>
        ''' An array filled with all the indices of the specified array.  The index of the returned array starts from 1.
        ''' </returns>
        Public Shared Function GetAllIndices(array1 As Primitive) As Primitive
            Return array1.GetAllIndices()
        End Function
        ''' <summary>
        ''' Gets the number of items stored in the array.
        ''' </summary>
        ''' <param name="array">
        ''' The array for which the count is requested.
        ''' </param>
        ''' <returns>
        ''' The number of items in the specified array.
        ''' </returns>
        Public Shared Function GetItemCount(array1 As Primitive) As Primitive
            Return array1.GetItemCount()
        End Function
        ''' <summary>
        ''' Gets whether or not a given variable is an array.
        ''' </summary>
        ''' <param name="array">
        ''' The variable to check.
        ''' </param>
        ''' <returns>
        ''' "True" if the specified variable is an array.  "False" otherwise.
        ''' </returns>
        Public Shared Function IsArray(array1 As Primitive) As Primitive
            Return array1.IsArray
        End Function
        ''' <summary>
        ''' Sets a value for a given array and index.
        ''' </summary>
        ''' <param name="arrayName">
        ''' The name of the array.
        ''' </param>
        ''' <param name="index">
        ''' Name of the index.
        ''' </param>
        ''' <param name="value">
        ''' The value to set.
        ''' </param>
        <HideFromIntellisense>
        Public Shared Sub SetValue(arrayName As Primitive, index As Primitive, value As Primitive)
            If Not arrayName.IsEmpty Then
                Dim value2 As Collections.Generic.Dictionary(Of Primitive, Primitive) = Nothing
                If Not _arrayMap.TryGetValue(arrayName, value2) Then
                    value2 = New Dictionary(Of Primitive, Primitive)
                    _arrayMap(arrayName) = value2
                End If
                value2(index) = value
            End If
        End Sub
        ''' <summary>
        ''' Gets a value for a given array and index.
        ''' </summary>
        ''' <param name="arrayName">
        ''' The name of the array.
        ''' </param>
        ''' <param name="index">
        ''' The name of the index.
        ''' </param>
        ''' <returns>
        ''' The value at the specified index of the specified array.
        ''' </returns>
        <HideFromIntellisense>
        Public Shared Function GetValue(arrayName As Primitive, index As Primitive) As Primitive
            Dim array1 As Dictionary(Of Primitive, Primitive) = GetArray(arrayName)
            If array1 Is Nothing Then
                Return ""
            End If

            Dim value As Primitive = Nothing
            If array1.TryGetValue(index, value) Then
                Return value
            End If

            Return ""
        End Function
        ''' <summary>
        ''' Removes the array item at the specified index.
        ''' </summary>
        ''' <param name="arrayName">
        ''' The name of the array.
        ''' </param>
        ''' <param name="index">
        ''' The index of the item to remove.
        ''' </param>
        <HideFromIntellisense>
        Public Shared Sub RemoveValue(arrayName As Primitive, index As Primitive)
            GetArray(arrayName)?.Remove(index)
        End Sub
        Friend Shared Function GetArray(arrayName As Primitive) As Dictionary(Of Primitive, Primitive)
            Dim value As Collections.Generic.Dictionary(Of Primitive, Primitive) = Nothing
            If _arrayMap.TryGetValue(arrayName, value) Then
                Return value
            End If

            Return Nothing
        End Function
    End Class
End Namespace
