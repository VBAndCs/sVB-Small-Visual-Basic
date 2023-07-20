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
    <SmallVisualBasicType>
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
                arr._arrayMap = New Dictionary(Of Primitive, Primitive)(Primitive.PrimitiveComparer.Instance)
                Return arr
            End Get
        End Property

        ''' <summary>
        ''' Adds an item to the array.
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <param name="value">the item you want to add after the last item in the array.</param>
        ''' <returns>a new array with the new item added. The input array will not be changed</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function AddItem(array As Primitive, value As Primitive) As Primitive
            Dim map = array._arrayMap
            Dim index = If(map Is Nothing, 1, map.Count + 1)
            Return Primitive.SetArrayValue(value, array, index)
        End Function

        ''' <summary>
        ''' Adds an item to the array. The input array will be changed directly, so this method is faster when you want to build a large array, but be careful because it will affect the reference array that the current array is copied from!
        ''' sVB calls this method when you use the array intializer {}, so it is faster than adding individual items using the array indixer [].
        ''' </summary>
        ''' <param name="array">The input array. The array must be intialized first, even with an empty array {}.</param>
        ''' <param name="value">The item you want to add after the last item in the array.</param>
        <WinForms.ReturnValueType(VariableType.Array)>
        <HideFromIntellisense>
        Public Shared Sub AddNextItem(array As Primitive, value As Primitive)
            Dim map = array._arrayMap
            If map Is Nothing Then
                Throw New InvalidOperationException("Can't add items to an empty array.")
            End If
            Dim index = map.Count + 1
            map(index) = value
        End Sub

        ''' <summary>
        ''' Adds an item at the end of the array. The input array will be changed directly, so this method is faster when you want to build a large array, but be careful because it will affect the reference array that the current array is copied from!
        ''' sVB calls this method when you use the array intializer {}, so it is faster than adding individual items using the array indixer [].
        ''' </summary>
        ''' <param name="array">The input array. The array must be intialized first, even with an empty array {}.</param>
        ''' <param name="value">The item you want to add after the last item in the array.</param>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Sub Append(array As Primitive, value As Primitive)
            Dim map = array._arrayMap
            If map Is Nothing Then
                Throw New InvalidOperationException("Can't add items to an empty array.")
            End If
            Dim index = map.Count + 1
            map(index) = value
        End Sub


        ''' <summary>
        ''' Adds many items to the array.
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <param name="items">an array containing the items to add eact of them as a single item at the end of the array.</param>
        ''' <returns>a new array with the given items. The input array will not be changed</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function AddItems(array As Primitive, items As Primitive) As Primitive
            If items.IsEmpty Then Return array
            If array.IsEmpty Then Return items

            Dim p As New Primitive
            Dim map As New Dictionary(Of Primitive, Primitive)(
                array._arrayMap,
                Primitive.PrimitiveComparer.Instance
            )

            Dim index As Integer = map.Count + 1
            p._arrayMap = map

            If items.IsArray Then
                For Each item In items._arrayMap.Values
                    map(Index) = item
                    Index += 1
                Next
            Else
                map(index) = items
            End If

            Return p
        End Function


        ''' <summary>
        ''' Adds an item to the array, with the given key and value
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <param name="key">the key of the item. If there is already an item with this key, it's value will be modified</param>
        ''' <param name="value">the value of the item</param>
        ''' <returns>a new arry with the item added. the input array will not be changed</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function AddKeyValue(
                   array As Primitive,
                   key As Primitive,
                   value As Primitive
        ) As Primitive

            Return Primitive.SetArrayValue(value, array, key)
        End Function

        ''' <summary>
        ''' Gets the value of the array item that exists at the given numeric position.
        ''' </summary>
        ''' <param name="array">The given array.</param>
        ''' <param name="position">A number that represents the position of the item in the array, which not always be the same as its index (key). Note that position must be > 0.</param>
        ''' <returns>the item value if found, otherwise empty string "".</returns>
        Public Shared Function GetItemAt(array As Primitive, position As Primitive) As Primitive
            If position < 1 Then Return ""

            Dim keys = array._arrayMap.Keys
            If position > keys.Count Then Return ""

            Return array._arrayMap(keys(position - 1))
        End Function

        ''' <summary>
        ''' Sets the value of the array item that exists at the given numeric position. The input array will be changed directly, so this method is faster then the array indexer [] when dealing with a large array, but be careful because it will affect the reference array that the current array is copied from!
        ''' </summary>
        ''' <param name="array">The given array.</param>
        ''' <param name="position">A number that represents the position of the item in the array, which is not always the same as its index (key). Note that position must be > 0.</param>
        ''' <param name="value">The item value.</param>
        ''' <returns>the item key if found at the given position, otherwise an empty string "".</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function SetItemAt(
                          array As Primitive,
                          position As Primitive,
                          value As Primitive
                   ) As Primitive

            If position < 1 Then Return ""

            Dim map = array._arrayMap
            If map Is Nothing OrElse map.Count = 0 Then Return ""

            Dim keys = map.Keys
            If position > keys.Count Then Return ""

            Dim key = keys(position - 1)
            map(key) = value

            Return key
        End Function

        ''' <summary>
        ''' Gets the key of the array item that exists at the given numeric position.
        ''' </summary>
        ''' <param name="array">The given array.</param>
        ''' <param name="position">A number that represents the position of the item in the array, which not always be the same as its index (key). Note that position must be > 0.</param>
        ''' <returns>the item key if found, otherwise empty string "".</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetKeyAt(array As Primitive, position As Primitive) As Primitive
            If position < 1 Then Return ""

            Dim keys = array._arrayMap.Keys
            If position > keys.Count Then Return ""

            Return keys(position - 1)
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
        ''' <returns>a new array with the value set. The input array will not be changed</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        <HideFromIntellisense>
        Public Shared Function SetValue(
                            array As Primitive,
                            index As Primitive,
                            value As Primitive
                    ) As Primitive

            Return Primitive.SetArrayValue(value, array, index)
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
        ''' <param name="index">The index of the item to remove.</param>
        ''' <returns>a new array with the item removed. The original array will not change</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function RemoveItem(array As Primitive, index As Primitive) As Primitive
            If index.AsString() = "" Then Return array

            Dim map = array._arrayMap
            If map Is Nothing OrElse Not map.ContainsKey(index) Then Return array

            Dim map2 As New Dictionary(Of Primitive, Primitive)(map, Primitive.PrimitiveComparer.Instance)
            map2?.Remove(index)

            Return New Primitive With {
                ._isArray = True,
                ._arrayMap = map2
            }
        End Function


        ''' <summary>
        ''' Searches the array for the given value, and return its key.
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <param name="value">the item to search for</param>
        ''' <param name="start">an integer representing the array index to start searching at</param>
        ''' <param name="ignoreCase">set it to true if you want to do an case-insensitive search</param>
        ''' <returns>the key of the item if found, otherwise empty string</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function Find(array As Primitive, value As Primitive, start As Primitive, ignoreCase As Primitive) As Primitive
            If array.IsEmpty OrElse value.IsEmpty Then Return ""
            Dim count = array._arrayMap.Count

            Dim intStart = System.Math.Max(CInt(start), 1)
            intStart = System.Math.Min(intStart, count)


            Dim ids = array._arrayMap.Keys
            Dim map = array._arrayMap
            Dim ignCase = CBool(ignoreCase)
            Dim lowercaseValue = value.AsString().ToLower()

            For i = intStart - 1 To count - 1
                If ignCase Then
                    If map(ids(i)).AsString().ToLower() = lowercaseValue Then
                        Return array._arrayMap.Keys(i)
                    End If
                ElseIf map(ids(i)) = value Then
                    Return array._arrayMap.Keys(i)
                End If
            Next

            Return ""
        End Function

        ''' <summary>
        ''' Searches the array for the given value, and returns its index position in the array.
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <param name="value">the item to search for</param>
        ''' <param name="start">an integer representing the array index to start searching at</param>
        ''' <param name="ignoreCase">set it to true if you want to do an case-insensitive search</param>
        ''' <returns>the position (order) of the item in the array if found, otherwise 0. </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function IndexOf(array As Primitive, value As Primitive, start As Primitive, ignoreCase As Primitive) As Primitive
            If array.IsEmpty OrElse value.IsEmpty Then Return ""
            Dim count = array._arrayMap.Count

            Dim intStart = System.Math.Max(CInt(start), 1)
            intStart = System.Math.Min(intStart, count)

            Dim ids = array._arrayMap.Keys
            Dim map = array._arrayMap
            Dim ignCase = CBool(ignoreCase)
            Dim lowercaseValue = value.AsString().ToLower()

            For i = intStart - 1 To count - 1
                If ignCase Then
                    If map(ids(i)).AsString().ToLower() = lowercaseValue Then
                        Return i
                    End If
                ElseIf map(ids(i)) = value Then
                    Return i
                End If
            Next

            Return 0
        End Function


        ''' <summary>
        ''' Joins the given array items into one text.
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <param name="separator">a string to use as a separator between array items</param>
        ''' <returns>a string containing the array items, separated by the given separator</returns>
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

        ''' <summary>
        '''Converts the given array to a string
        ''' </summary>
        ''' <param name="array">the input array</param>
        ''' <returns>The string representation of the array. If the array is simple, its string can be of the form {1, 2, 3}. If the array contains keys, its string can be of the form {[1]=1, [2]=2, [Name]=Adam}.</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function ToStr(array As Primitive) As Primitive
            Return Text.ToStr(array)
        End Function
    End Class
End Namespace
