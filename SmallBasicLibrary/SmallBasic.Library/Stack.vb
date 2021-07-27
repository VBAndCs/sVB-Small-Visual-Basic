
Namespace Library
    ''' <summary>
    ''' This object provides a way of storing values just like stacking up a plate.  You can push a value to the top of the stack and pop it off. You can only pop the values one by one off the stack and the last pushed value will be the first one to pop out.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Stack
        Private Shared _stackMap As New Dictionary(Of Primitive, Stack(Of Primitive))

        ''' <summary>
        ''' Pushes a value to the specified stack.
        ''' </summary>
        ''' <param name="stackName">
        ''' The name of the stack.
        ''' </param>
        ''' <param name="value">
        ''' The value to push.
        ''' </param>
        Public Shared Sub PushValue(stackName As Primitive, value As Primitive)
            Dim value2 As Stack(Of Primitive) = Nothing
            If Not _stackMap.TryGetValue(stackName, value2) Then
                value2 = New Stack(Of Primitive)
                _stackMap(stackName) = value2
            End If
            value2.Push(value)
        End Sub

        ''' <summary>
        ''' Gets the count of items in the specified stack.
        ''' </summary>
        ''' <param name="stackName">
        ''' The name of the stack.
        ''' </param>
        ''' <returns>
        ''' The number of items in the specified stack.
        ''' </returns>
        Public Shared Function GetCount(stackName As Primitive) As Primitive
            Dim value As Collections.Generic.Stack(Of Primitive) = Nothing
            If Not _stackMap.TryGetValue(stackName, value) Then
                value = New Stack(Of Primitive)
                _stackMap(stackName) = value
            End If

            Return value.Count
        End Function

        ''' <summary>
        ''' Pops the value from the specified stack.
        ''' </summary>
        ''' <param name="stackName">
        ''' The name of the stack.
        ''' </param>
        ''' <returns>
        ''' The value from the stack.
        ''' </returns>
        Public Shared Function PopValue(stackName As Primitive) As Primitive
            Dim value As Collections.Generic.Stack(Of Primitive) = Nothing
            If _stackMap.TryGetValue(stackName, value) Then
                Return value.Pop()
            End If

            Return ""
        End Function
    End Class
End Namespace
