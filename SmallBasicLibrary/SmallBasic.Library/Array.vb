' To configure or remove Option's included in result, go to Options/Advanced Options...
Option Compare Text
Option Explicit On
Option Infer On
Option Strict On
Imports System.Collections.Generic

Namespace Library

    ''' <summary>
    ''' This object provides a way of storing more than one value for a given name. These values can be accessed by another index.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Array

        ''' <summary>
        ''' Creates an empty array. You can also use x = {} to create an empty array.
        ''' </summary>
        ''' <returns>an empty array</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared ReadOnly Property EmptyArray As Primitive
            Get
                Dim arr As Primitive = ""
                arr._isArray = True
                Return arr
            End Get
        End Property

        ''' <summary>
        ''' Adds an item to the array.
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <param name="value">the item you want to add after the last item in the array.</param>
        ''' <returns>a new array with the new item added. The input array will bot be cahmged</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function AddItem(array As Primitive, value As Primitive) As Primitive
            Dim map = array._arrayMap
            If map Is Nothing Then
                Return SetValue(array, 1, value)
            Else
                Return SetValue(array, map.Count + 1, value)
            End If
        End Function

        ''' <summary>
        ''' Adds many items to the array.
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <param name="items">an array containing the items to add eact of them as a single item at the end of the array.</param>
        ''' <returns>a new array with the new items added. The input array will bot be cahmged</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function AddItems(array As Primitive, items As Primitive) As Primitive
            If items.IsEmpty Then Return array
            If items.IsArray Then
                Dim index As Integer = array.GetItemCount() + 1
                For Each item In items._arrayMap.Values
                    array(index) = item
                    index += 1
                Next
                Return array
            Else
                Return AddItem(array, items)
            End If
        End Function


        ''' <summary>
        ''' Adds an item to the array, with the given key and value
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <param name="key">the key of the item. If there is already an item with this key, it's value will be modified</param>
        ''' <param name="value">the value of the item</param>
        ''' <returns>a new arry with the item added. the input array will not be changed</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function AddKeyValue(array As Primitive, key As Primitive, value As Primitive) As Primitive
            Return SetValue(array, key, value)
        End Function

        <HideFromIntellisense>
        Public Shared Function GetItemAt(array As Primitive, index As Primitive) As Primitive
            Return array(index)
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="array">the input array.</param>
        ''' <param name="index">the index or the key of the item you want to change</param>
        ''' <param name="value">the value of the item</param>
        ''' <returns>a new array with the value set. The input array will bot be cahmged</returns>
        <HideFromIntellisense>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function SetItemAt(
                          array As Primitive,
                          index As Primitive,
                          value As Primitive
                   ) As Primitive

            array(index) = value
            Return array
        End Function

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
        '''True" or "False" depending on if the index was present in the specified array.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function ContainsIndex(array As Primitive, index As Primitive) As Primitive
            Return array.ContainsKey(index)
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
        ''' "True" or "False" depending on if the value was present in the specified array.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function ContainsValue(array As Primitive, value As Primitive) As Primitive
            Return array.ContainsValue(value)
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
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function GetAllIndices(array As Primitive) As Primitive
            Return array.GetAllIndices()
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
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetItemCount(array As Primitive) As Primitive
            Return array.GetItemCount()
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
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function IsArray(array As Primitive) As Primitive
            Return array.IsArray
        End Function

        ''' <summary>
        ''' Sets a value for a given array and index.
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <param name="index">the index or the key of the item.</param>
        ''' <param name="value">the value to set.''' </param>
        ''' <returns>a new array with the value set. The input array will bot be cahmged</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        <HideFromIntellisense>
        Public Shared Function SetValue(
                            array As Primitive,
                            index As Primitive,
                            value As Primitive
                    ) As Primitive

            array(index) = value
            Return array
        End Function

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
            Dim map = arrayName._arrayMap
            If map Is Nothing OrElse Not map.ContainsKey(index) Then
                Return ""
            End If

            Return map(index)
        End Function

        ''' <summary>
        ''' Removes the array item at the specified index.
        ''' </summary>
        ''' <param name="array">The name of the array.</param>
        ''' <param name="index">
        ''' The index of the item to remove.
        ''' </param>
        Public Shared Sub RemoveItemAt(array As Primitive, index As Primitive)
            array._arrayMap?.Remove(index)
        End Sub

        ''' <summary>
        ''' Removes all items from the array.
        ''' </summary>
        ''' <param name="array">The name of the array.</param>
        Public Shared Sub Clear(array As Primitive)
            array._arrayMap?.Clear()
        End Sub

        ''' <summary>
        ''' Searches the array for the given value.
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <param name="value">the item to sarch for</param>
        ''' <param name="start">an integer representing the array index to sratr swarching at</param>
        ''' <param name="ignoreCase">set it to true if you want to do an case-insensetive search</param>
        ''' <returns>the key of the item if found, otherwise empty string</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function Find(array As Primitive, value As Primitive, start As Primitive, ignoreCase As Primitive) As Primitive
            If array.IsEmpty OrElse value.IsEmpty Then Return ""
            Dim count = array._arrayMap.Count

            Dim intStart = System.Math.Max(CInt(start), 1)
            intStart = System.Math.Min(intStart, count)

            Dim values = array._arrayMap.Values()
            Dim ignCase = CBool(ignoreCase)
            Dim lowercaseValue = value.AsString().ToLower()

            For i = intStart To count
                If ignCase Then
                    If values(i).AsString().ToLower() = lowercaseValue Then
                        Return array._arrayMap.Keys(i)
                    End If
                ElseIf values(i) = value Then
                    Return array._arrayMap.Keys(i)
                End If
            Next

            Return ""
        End Function

        ''' <summary>
        ''' Joins the given array items into one text.
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <param name="separator">a string to use as a separator between array items</param>
        ''' <returns>a string containing the array items, seperated by the given separator</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function Join(array As Primitive, separator As Primitive) As Primitive
            If array.IsEmpty Then Return ""
            If array.IsArray Then
                Dim sb As New System.Text.StringBuilder
                Dim arr = array._arrayMap.Values
                Dim n = arr.Count - 1
                Dim sep = separator.AsString()
                For i = 0 To n
                    sb.Append(arr(i).AsString())
                    If i < n Then sb.Append(sep)
                Next
                Return sb.ToString()
            Else
                Return array
            End If
        End Function

    End Class
End Namespace
