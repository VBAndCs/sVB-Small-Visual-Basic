Imports System.Text
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.Expressions
    <Serializable>
    Public Class InitializerExpression
        Inherits Expression
        Public ReadOnly Arguments As New List(Of Expression)

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(precedence As Integer, arguments As List(Of Expression))
            Me.Precedence = precedence
            Me.Arguments = arguments
        End Sub


        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)

            For Each argument In Arguments
                argument.Parent = Me.Parent
                argument.AddSymbols(symbolTable)
            Next
        End Sub

        Private Shared counter = 0

        Private Function Lower(leftValue As String, Optional ByRef tmpArr As String = Nothing) As String
            counter += 1
            tmpArr = "__tmpArray__" & counter

            If Arguments.Count = 0 Then
                If leftValue = "" Then leftValue = tmpArr
                Return $"{leftValue} = Array.EmptyArray"
            End If

            Dim code As New StringBuilder($"{tmpArr} = """"")
            code.AppendLine()

            If TypeOf Arguments(0) Is InitializerExpression Then
                Dim initExpr = CType(Arguments(0), InitializerExpression)
                code.AppendLine()
                code.AppendLine(initExpr.Lower($"{tmpArr}[1]"))
            Else
                code.AppendLine($"{tmpArr}[{1}] = {Arguments(0)}")
            End If

            For i = 1 To Arguments.Count - 1
                If TypeOf Arguments(i) Is InitializerExpression Then
                    Dim initExpr = CType(Arguments(i), InitializerExpression)
                    code.AppendLine()
                    Dim tmpArr2 As String
                    code.AppendLine(initExpr.Lower("", tmpArr2))
                    code.AppendLine($"Array.AddNextItem({tmpArr}, {tmpArr2})")
                Else
                    code.AppendLine($"Array.AddNextItem({tmpArr}, {Arguments(i)})")
                End If
            Next

            If leftValue <> "" Then
                code.AppendLine($"{leftValue} = {tmpArr}")
            End If

            Return code.ToString()
        End Function

        Public Function LowerAndEmit(
                        leftValue As String,
                        scope As CodeGenScope,
                        lineOffset As Integer
                    ) As Expression

            Dim code = Me.Lower(leftValue)
            Dim subroutine = SubroutineStatement.GetSubroutine(Me)
            If subroutine Is Nothing Then subroutine = SubroutineStatement.Current
            Return Parser.ParseAndEmit(code, subroutine, scope, lineOffset)
        End Function


        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp Then
                ' no need to emit the array. This is just for global file help,
                ' and we know what we need about the array from the symbol table
                LiteralExpression.Zero.EmitIL(scope)
            Else
                ' Lefthand value must be an empty string
                LowerAndEmit("", scope, StartToken.Line).EmitIL(scope)
            End If
        End Sub

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder("{")

            If Arguments.Count > 0 Then
                For Each arg In Arguments
                    sb.Append(arg.ToString())
                    sb.Append(", ")
                Next

                sb.Remove(sb.Length - 2, 2)
            End If

            sb.Append("}")
            Return sb.ToString()
        End Function

        Public Overrides Function InferType(symbolTable As SymbolTable) As VariableType
            Return VariableType.Array
        End Function

        Public Overrides Function Evaluate(runner As Engine.ProgramRunner) As Primitive
            Dim count = Arguments.Count
            If count = 0 Then Return Array.EmptyArray()

            Dim arr As New Primitive()
            arr.Items(1) = Arguments(0).Evaluate(runner)
            For i = 1 To count - 1
                Array.Append(arr, Arguments(i).Evaluate(runner))
            Next
            Return arr
        End Function
    End Class
End Namespace
