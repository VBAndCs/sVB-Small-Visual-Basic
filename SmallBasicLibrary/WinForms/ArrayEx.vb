' To configure or remove Option's included in result, go to Options/Advanced Options...
Option Compare Text
Option Explicit On
Option Infer On
Option Strict On
Imports System.Collections.Generic
Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class ArrayEx

        ''' <summary>
        ''' Adds an item to the array.
        ''' </summary>
        ''' <param name="value">the item you want to add after the last item in the array.</param>
        ''' <returns>a new array with the new item added. The input array will bot be cahmged</returns>
        <WinForms.ExMethod>
        <ReturnValueType(VariableType.Array)>
        Public Shared Function AddItem(array As Primitive, value As Primitive) As Primitive
            Return Library.Array.AddItem(array, value)
        End Function

        ''' <summary>
        ''' Adds many items to the array.
        ''' </summary>
        ''' <param name="items">an array containing the items to add eact of them as a single item at the end of the array.</param>
        ''' <returns>a new array with the new items added. The input array will bot be cahmged</returns>
        <WinForms.ExMethod>
        <ReturnValueType(VariableType.Array)>
        Public Shared Function AddItems(array As Primitive, items As Primitive) As Primitive
            Return Library.Array.AddItems(array, items)
        End Function


        ''' <summary>
        ''' Adds an item to the array, with the given key and value
        ''' </summary>
        ''' <param name="key">the key of the item. If there is already an item with this key, it's value will be modified</param>
        ''' <param name="value">the value of the item</param>
        ''' <returns>a new arry with the item added. the input array will not be changed</returns>
        <WinForms.ExMethod>
        <ReturnValueType(VariableType.Array)>
        Public Shared Function AddKeyValue(array As Primitive, key As Primitive, value As Primitive) As Primitive
            Return Library.Array.SetValue(array, key, value)
        End Function

        ''' <summary>
        ''' Gets whether or not the array contains the specified index.  This is very useful when deciding if the array's index was initialized by some value or not.
        ''' </summary>
        ''' <param name="index"> The index to check.</param>
        ''' <returns>
        ''' "True" or "False" depending on if the index was present in the specified array.
        ''' </returns>
        <WinForms.ExMethod>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function ContainsIndex(array As Primitive, index As Primitive) As Primitive
            Return Library.Array.ContainsIndex(array, index)
        End Function

        ''' <summary>
        ''' Gets whether or not the array contains the specified value.  This is very useful when deciding if the array's value was stored in some index.
        ''' </summary>
        ''' <param name="value">The value to check.</param>
        ''' <returns>
        ''' "True" or "False" depending on if the value was present in the specified array.
        ''' </returns>
        <WinForms.ExMethod>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function ContainsValue(array As Primitive, value As Primitive) As Primitive
            Return Library.Array.ContainsValue(array, value)
        End Function

        ''' <summary>
        ''' Gets all the indices for the array, as another array.
        ''' </summary>
        ''' <returns>
        ''' An array filled with all the indices of the specified array.  The index of the returned array starts from 1.
        ''' </returns>
        <WinForms.ExProperty>
        <ReturnValueType(VariableType.Array)>
        Public Shared Function GetIndices(array As Primitive) As Primitive
            Return array.GetAllIndices()
        End Function

        ''' <summary>
        ''' Gets the number of items stored in the array.
        ''' </summary>
        ''' <returns>
        ''' The number of items in the specified array.
        ''' </returns>
        <WinForms.ExProperty>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetCount(array As Primitive) As Primitive
            Return array.GetItemCount()
        End Function

        ''' <summary>
        ''' Removes the array item at the specified index.
        ''' </summary>
        ''' <param name="index">
        ''' The index of the item to remove.
        ''' </param>
        <WinForms.ExMethod>
        Public Shared Sub RemoveAt(array As Primitive, index As Primitive)
            array._arrayMap?.Remove(index)
        End Sub

        ''' <summary>
        ''' Removes all items from the array.
        ''' </summary>
        <WinForms.ExMethod>
        Public Shared Sub Clear(array As Primitive)
            array._arrayMap?.Clear()
        End Sub

        ''' <summary>
        ''' Searches the current array for the given value.
        ''' </summary>
        ''' <param name="value">the item to sarch for</param>
        ''' <param name="start">an integer representing the array index to sratr swarching at</param>
        ''' <param name="ignoreCase">set it to true if you want to do an case-insensetive search</param>
        ''' <returns>the index or string key of the item if found, otherwise empty string</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function Find(
                        array As Primitive,
                        value As Primitive,
                        start As Primitive,
                        ignoreCase As Primitive
                  ) As Primitive

            Return Library.Array.Find(array, value, start, ignoreCase)
        End Function

        ''' <summary>
        ''' Joins the given array items into one text.
        ''' </summary>
        ''' <param name="separator">a string to use as a separator between array items</param>
        ''' <returns>a string containing the array items, seperated by the given separator</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function Join(array As Primitive, separator As Primitive) As Primitive
            Return Library.Array.Join(array, separator)
        End Function

    End Class
End Namespace
