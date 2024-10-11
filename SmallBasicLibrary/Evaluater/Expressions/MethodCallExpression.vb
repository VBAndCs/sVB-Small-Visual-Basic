Imports System.Text
Imports Microsoft.SmallVisualBasic.Library

Namespace Evaluator.Expressions
    <Serializable>
    Friend Class MethodCallExpression
        Inherits Expression

        Public Shared MathFunctions() As System.Reflection.MethodInfo

        Public Sub New(startToken As Token, precedence As Integer, methodName As Token, arguments As List(Of Expression))
            Me.StartToken = startToken
            Me.Precedence = precedence
            _MethodName = methodName
            _Arguments = arguments

            If MathFunctions Is Nothing Then
                MathFunctions = GetType(Library.Math).GetMethods(System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
            End If
        End Sub

        Public Property MethodName As Token

        Public ReadOnly Property Arguments As New List(Of Expression)

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder()
            sb.Append(MethodName.Text & "(")

            If Arguments.Count > 0 Then
                For Each arg In Arguments
                    sb.Append(arg.ToString())
                    sb.Append(", ")
                Next

                sb.Remove(sb.Length - 2, 2)
            End If

            sb.Append(")")
            Return sb.ToString()
        End Function

        Public Overrides Function Evaluate(x As Double) As Library.Primitive
            Dim name = _MethodName.LCaseText
            Select Case name
                Case "rnd", "rand", "random", "getrnd", "getrand", "getrandom"
                    name = "getrandomnumber"
                Case "rad", "radians"
                    name = "getradians"
                Case "deg", "degrees"
                    name = "getdegrees"
                Case "sqrt"
                    name = "squareroot"
                Case "asin"
                    name = "arcsin"
                Case "acos"
                    name = "arccos"
                Case "atan"
                    name = "arctan"
                Case "pow"
                    name = "power"
                Case "round"
                    If Arguments.Count = 2 Then
                        name = "round2"
                    End If
                Case "ln"
                    name = "naturallog"
            End Select

            Dim methods = From m In MathFunctions
                          Where m.Name.ToLower() = name

            If methods.Any Then
                Dim method = methods.First
                Dim paramsCount = method.GetParameters().Length
                If paramsCount = Arguments.Count Then
                    Dim n = Arguments.Count - 1
                    Dim args(n) As Object
                    For i = 0 To n
                        args(i) = Arguments(i).Evaluate(x)
                    Next
                    Return CDbl(CType(method.Invoke(Nothing, args), Library.Primitive))

                Else
                    Evaluator.Errors.Add(New [Error](
                         _MethodName,
                         $"Math.{method.Name} has {paramsCount} parameters, but you sent {Arguments.Count} arguments."
                    ))
                End If

            Else
                Evaluator.Errors.Add(New [Error](
                     _MethodName,
                     $"Unrecognized method name {_MethodName.Text}"
                ))
            End If

            Return New Primitive("")
        End Function
    End Class
End Namespace
