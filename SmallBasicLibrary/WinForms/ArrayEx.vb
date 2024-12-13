' To configure or remove Option's included in result, go to Options/Advanced Options...
Option Compare Text
Option Explicit On
Option Infer On
Option Strict On
Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class ArrayEx

        ''' <summary>
        ''' Adds an item to the array.
        ''' </summary>
        ''' <param name="value">the item you want to add after the last item in the array.</param>
        ''' <returns>a new array with the new item added. The input array will not be changed</returns>
        <WinForms.ExMethod>
        <ReturnValueType(VariableType.Array)>
        Public Shared Function AddItem(array As Primitive, value As Primitive) As Primitive
            Return Library.Array.AddItem(array, value)
        End Function

        ''' <summary>
        ''' Adds many items to the array.
        ''' </summary>
        ''' <param name="items">an array containing the items to add eact of them as a single item at the end of the array.</param>
        ''' <returns>a new array with the new items added. The input array will not be changed</returns>
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
        ''' <returns>a new array with the item added. the input array will not be changed</returns>
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
        ''' Removes the array item at the specified index (key).
        ''' </summary>
        ''' <param name="index">The key of the item to remove.</param>
        ''' <returns>a new array with the item removed. The original array will not change</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        <ExMethod>
        Public Shared Function RemoveItem(array As Primitive, index As Primitive) As Primitive
            Return Library.Array.RemoveItem(array, index)
        End Function



        ''' <summary>
        ''' Searches the array for the given value, and returns its index position in the array.
        ''' </summary>
        ''' <param name="value">the item to search for</param>
        ''' <param name="start">an integer representing the array index to start searching at</param>
        ''' <param name="ignoreCase">set it to true if you want to do an case-insensitive search</param>
        ''' <returns>the position (order) of the item in the array if found, otherwise 0. </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        <WinForms.ExMethod>
        Public Shared Function IndexOf(array As Primitive, value As Primitive, start As Primitive, ignoreCase As Primitive) As Primitive
            Return Library.Array.IndexOf(array, value, start, ignoreCase)
        End Function

        ''' <summary>
        ''' Gets value of the array item that exists at the given numeric position.
        ''' </summary>
        ''' <param name="position">A number that represents the position of the item in the array, which not always be the same as its index (key). Note that position must be > 0.</param>
        ''' <returns>the item value if found, otherwise empty string "".</returns>
        <ExMethod>
        Public Shared Function GetItemAt(array As Primitive, position As Primitive) As Primitive
            Return Library.Array.GetItemAt(array, position)
        End Function

        ''' <summary>
        ''' Sets the value of the array item that exists at the given numeric position. The current array will be changed directly, so this method is faster then the array indexer [] when dealing with a large array, but be careful because it will affect the reference array that the current array is copied from!
        ''' </summary>
        ''' <param name="position">A number that represents the position of the item in the array, which is not always the same as its index (key). Note that position must be > 0.</param>
        ''' <param name="value">The item value.</param>
        ''' <returns>the item key if found at the given position, otherwise an empty string "".</returns>
        <ExMethod>
        <ReturnValueType(VariableType.String)>
        Public Shared Function SetItemAt(
                          array As Primitive,
                          position As Primitive,
                          value As Primitive
                   ) As Primitive
            Return Library.Array.SetItemAt(array, position, value)
        End Function

        ''' <summary>
        ''' Gets the key of the array item that exists at the given numeric position.
        ''' </summary>
        ''' <param name="position">A number that represents the position of the item in the array, which not always be the same as its index (key). Note that position must be > 0.</param>
        ''' <returns>the item key if found, otherwise empty string "".</returns>
        <ExMethod>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetKeyAt(array As Primitive, position As Primitive) As Primitive
            Return Library.Array.GetKeyAt(array, position)
        End Function

        ''' <summary>
        ''' Searches the current array for the given value.
        ''' </summary>
        ''' <param name="value">the item to search for</param>
        ''' <param name="start">an integer representing the array index to start searching at</param>
        ''' <param name="ignoreCase">set it to true if you want to do an case-insensetive search</param>
        ''' <returns>the index or string key of the item if found, otherwise an empty string</returns>
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
        ''' <returns>a string containing the array items, separated by the given separator</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function Join(array As Primitive, separator As Primitive) As Primitive
            Return Library.Array.Join(array, separator)
        End Function


        ''' <summary>
        '''Converts the current array to a string.
        ''' </summary>
        ''' <returns>The string representation of the array. If the array is simple, its string can be of the form {1, 2, 3}. If the array contains keys, its string can be of the form {[1]=1, [2]=2, [Name]=Adam}.</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function ToStr(array As Primitive) As Primitive
            Return Text.ToStr(array)
        End Function


        ''' <summary>
        ''' Adds an item to the array. The current array will be changed directly, so this method is faster when you want to build a large array, but be careful because it will affect the reference array that the current array is copied from!
        ''' sVB calls this method when you use the array intializer {}, so it is faster than adding individual items using the array indixer [].
        ''' </summary>
        ''' <param name="value">The item you want to add after the last item in the array.</param>
        <ExMethod>
        <HideFromIntellisense>
        Public Shared Sub AddNextItem(array As Primitive, value As Primitive)
            Library.Array.Append(array, value)
        End Sub

        ''' <summary>
        ''' Adds an item at the end of the array. The current array will be changed directly, so this method is faster when you want to build a large array, but be careful because it will affect the reference array that the current array is copied from!
        ''' sVB calls this method when you use the array intializer {}, so it is faster than adding individual items using the array indixer [].
        ''' </summary>
        ''' <param name="value">The item you want to add after the last item in the array.</param>
        <ExMethod>
        Public Shared Sub Append(array As Primitive, value As Primitive)
            Library.Array.Append(array, value)
        End Sub
    End Class
End Namespace
