Imports System.Globalization
Imports System.Text
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.Expressions
    <Serializable>
    Public Class InitializerExpression
        Inherits Expression
        Public ReadOnly Property Arguments As New List(Of Expression)

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(precedence As Integer, arguments As List(Of Expression))
            Me.Precedence = precedence
            _Arguments = arguments
        End Sub


        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)

            For Each argument In Arguments
                argument.Parent = Me.Parent
                argument.AddSymbols(symbolTable)
            Next
        End Sub

        Private Shared counter = 0

        Private Function Lower(leftValue As String) As String
            If Arguments.Count = 0 Then Return $"{leftValue} = Array.EmptyArray"
            counter += 1
            Dim tmpVar = "__tmpArray__" & counter
            Dim code As New StringBuilder($"{tmpVar} = """"")
            code.AppendLine()

            For i = 0 To Arguments.Count - 1
                If TypeOf Arguments(i) Is InitializerExpression Then
                    Dim initExpr = CType(Arguments(i), InitializerExpression)
                    code.AppendLine()
                    code.AppendLine(initExpr.Lower($"{tmpVar}[{i + 1}]"))
                Else
                    code.AppendLine($"{tmpVar}[{i + 1}] = {Arguments(i)}")
                End If
            Next

            If leftValue <> "" Then
                code.AppendLine($"{leftValue} = {tmpVar}")
            End If
            counter -= 1
            Return code.ToString()
        End Function

        Public Function LowerAndEmit(leftValue As String, scope As CodeGenScope, lineOffset As Integer) As Expression
            Dim code = Me.Lower(leftValue)
            Dim subroutine = SubroutineStatement.GetSubroutine(Me)
            If subroutine Is Nothing Then subroutine = SubroutineStatement.Current
            Return ArrayExpression.ParseAndEmit(code, subroutine, scope, lineOffset)
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
            Dim stringBuilder As New StringBuilder("{")

            If Arguments.Count > 0 Then
                For Each argument In Arguments
                    stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}, ", New Object(0) {argument})
                Next

                stringBuilder.Remove(stringBuilder.Length - 2, 2)
            End If

            stringBuilder.Append("}")
            Return stringBuilder.ToString()
        End Function
    End Class
End Namespace
