
Namespace Library
    ''' <summary>
    ''' This object provides a way of storing values just like stacking up a plate.  You can push a value to the top of the stack and pop it off. You can only pop the values one by one off the stack and the last pushed value will be the first one to pop out.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Stack
        Private Shared _stackMap As New Dictionary(Of Primitive, Stack(Of Primitive))

        ''' <summary>
        ''' Pushes a value to the specified stack.
        ''' </summary>
        ''' <param name="stackName">The name of the stack.</param>
        ''' <param name="value">The value to push.</param>
        Public Shared Sub PushValue(stackName As Primitive, value As Primitive)
            Dim stack As Stack(Of Primitive) = Nothing
            stackName = stackName.ToString().ToLower()
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
            stackName = stackName.ToString().ToLower()
            If Not _stackMap.TryGetValue(stackName, stack) Then
                stack = New Stack(Of Primitive)
                _stackMap(stackName) = stack
            End If

            Return stack.Count
        End Function

        ''' <summary>
        ''' Pops the value from the specified stack.
        ''' </summary>
        ''' <param name="stackName">The name of the stack.</param>
        ''' <returns>
        ''' The value from the stack.
        ''' </returns>
        Public Shared Function PopValue(stackName As Primitive) As Primitive
            Dim stack As Stack(Of Primitive) = Nothing
            If _stackMap.TryGetValue(stackName.ToString().ToLower(), stack) Then
                Return stack.Pop()
            End If

            Return ""
        End Function
    End Class
End Namespace
