
Namespace Library
    ''' <summary>
    ''' This object provides a way of storing values just like stacking up a plate.  You can push a value to the top of the stack and pop it off. You can only pop the values one by one off the stack and the last pushed value will be the first one to pop out.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Stack
        Friend Shared _stackMap As New Dictionary(Of Primitive, Stack(Of Primitive))

        ''' <summary>
        ''' Pushes a value to the specified stack.
        ''' </summary>
        ''' <param name="stackName">The name of the stack.</param>
        ''' <param name="value">The value to push.</param>
        Public Shared Sub PushValue(stackName As Primitive, value As Primitive)
            Dim stack As Stack(Of Primitive) = Nothing
            stackName = Text.ToLower(stackName)
            If Not _stackMap.TryGetValue(stackName, stack) Then
                stack = New Stack(Of Primitive)
                _stackMap(stackName) = stack
            End If
            stack.Push(value)
        End Sub

        ''' <summary>
        ''' Gets the count of items in the specified stack.
        ''' </summary>
        ''' <param name="stackName">The name of the stack.</param>
        ''' <returns>
        ''' The number of items in the specified stack.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetCount(stackName As Primitive) As Primitive
            Dim stack As Stack(Of Primitive) = Nothing
            stackName = Text.ToLower(stackName)
            If Not _stackMap.TryGetValue(stackName, stack) Then
                stack = New Stack(Of Primitive)
                _stackMap(stackName) = stack
            End If

            Return stack.Count
        End Function

        ''' <summary>
        ''' Removees all items from the specified stack.
        ''' </summary>
        ''' <param name="stackName">The name of the stack.</param>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Sub Clear(stackName As Primitive)
            Dim stack As Stack(Of Primitive) = Nothing
            stackName = Text.ToLower(stackName)
            If _stackMap.TryGetValue(stackName, stack) Then
                stack.Clear()
            End If
        End Sub


        ''' <summary>
        ''' Pops the value from the specified stack.
        ''' </summary>
        ''' <param name="stackName">The name of the stack.</param>
        ''' <returns>
        ''' The value from the stack.
        ''' </returns>
        Public Shared Function PopValue(stackName As Primitive) As Primitive
            Dim stack As Stack(Of Primitive) = Nothing
            If _stackMap.TryGetValue(Text.ToLower(stackName), stack) Then
                If stack.Count > 0 Then Return stack.Pop()
            End If
            Return New Primitive("")
        End Function

        ''' <summary>
        ''' reads the last value pushed into the specified stack, whithout poping it out.
        ''' </summary>
        ''' <param name="stackName">The name of the stack.</param>
        ''' <returns>
        ''' The top value of the stack.
        ''' </returns>
        Public Shared Function PeekValue(stackName As Primitive) As Primitive
            Dim stack As Stack(Of Primitive) = Nothing
            If _stackMap.TryGetValue(Text.ToLower(stackName), stack) Then
                If stack.Count > 0 Then Return stack.Peek()
            End If
            Return New Primitive("")
        End Function

    End Class
End Namespace
