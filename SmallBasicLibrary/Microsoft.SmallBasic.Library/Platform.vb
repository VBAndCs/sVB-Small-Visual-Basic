' To configure or remove Option's included in result, go to Options/Advanced Options...
Option Compare Text
Option Explicit On
Option Infer On
Option Strict On
Imports System.Collections.Generic
Imports System.Reflection

Namespace Microsoft.SmallBasic.Library
    ''' <summary>
    ''' The Platform object provides a way to generically invoke other .Net libraries.
    ''' </summary>
    Public NotInheritable Class Platform
        Private Shared _nameGenerationMap As New Dictionary(Of String, Integer)
        Private Shared _objectMap As New Dictionary(Of String, Object)
        Public Shared Function CreateInstance(typeName As Primitive) As Primitive
            Try
                Dim type1 As Type = Type.[GetType](typeName)
                Dim value As Object = Activator.CreateInstance(type1)
                Dim text As String = GenerateNewName("Instance")
                _objectMap(text) = value
                Return text
            Catch
                Return "ERROR"
            End Try
        End Function
        Public Shared Function InvokeInstanceMethod(instanceId As Primitive, methodName As Primitive, argumentsStackName As Primitive) As Primitive
            Try
                Dim obj As Object = _objectMap(instanceId)
                Dim method As MethodInfo = obj.[GetType]().GetMethod(methodName, BindingFlags.IgnoreCase Or BindingFlags.Instance Or BindingFlags.[Public])
                Dim num As Integer = Stack.GetCount(argumentsStackName)
                Dim parameters As ParameterInfo() = method.GetParameters()
                If parameters.Length = num Then
                    Dim list1 As New List(Of Object)
                    For num2 As Integer = num - 1 To 0 Step -1
                        list1(num2) = Stack.PopValue(argumentsStackName)
                    Next

                    Return New Primitive(method.Invoke(obj, list1.ToArray()))
                End If

                Return "ERROR"
            Catch
                Return "ERROR"
            End Try
        End Function
        Public Shared Function InvokeStaticMethod(typeName As Primitive, methodName As Primitive, argumentsStackName As Primitive) As Primitive
            Try
                Dim type1 As Type = Type.[GetType](typeName)
                Dim method As MethodInfo = type1.GetMethod(methodName, BindingFlags.IgnoreCase Or BindingFlags.[Static] Or BindingFlags.[Public])
                Dim num As Integer = Stack.GetCount(argumentsStackName)
                Dim parameters As ParameterInfo() = method.GetParameters()
                If parameters.Length = num Then
                    Dim list1 As New List(Of Object)
                    For num2 As Integer = num - 1 To 0 Step -1
                        list1(num2) = Stack.PopValue(argumentsStackName)
                    Next

                    Return New Primitive(method.Invoke(Nothing, list1.ToArray()))
                End If

                Return "ERROR"
            Catch
                Return "ERROR"
            End Try
        End Function
        Friend Shared Function GenerateNewName(prefix As String) As String
            Dim value As Integer = 0
            _nameGenerationMap.TryGetValue(prefix, value)
            value += 1
            _nameGenerationMap(prefix) = value
            Return prefix & value
        End Function
    End Class
End Namespace
