Imports System.Globalization
Imports System.Reflection.Emit
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Statements
    Public Class SubroutineCallStatement
        Inherits Statement

        Public Name As TokenInfo
        Public Args As List(Of Expression)
        Public IsFunctionCall As Boolean
        Public OuterSubRoutine As SubroutineStatement

        Public Overrides Function GetStatement(lineNumber) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If Args.Count > 0 AndAlso lineNumber <= Args.Last.EndToken.Line Then Return Me
            Return Nothing
        End Function

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            Name.Parent = Me

            For Each arg In Args
                arg.Parent = Me
                arg.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If Args.Count > 0 Then
                Dim code As New Text.StringBuilder()
                For i = Args.Count - 1 To 0 Step -1
                    code.AppendLine($"Stack.PushValue(""_sVB_Arguments"", {Args(i)})")
                Next

                CodeGenerator.LowerAndEmit(code.ToString(), scope, OuterSubRoutine)
            End If

            Dim methodInfo = scope.MethodBuilders(Name.NormalizedText)
            scope.ILGenerator.EmitCall(OpCodes.Call, methodInfo, Nothing)

            If IsFunctionCall Then
                MethodCallStatement.DoNotPopReturnValue = True
                CodeGenerator.LowerAndEmit($"Stack.PopValue(""_sVB_ReturnValues"")", scope, OuterSubRoutine)
                MethodCallStatement.DoNotPopReturnValue = False
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0}()", New Object(0) {Name.Text})
        End Function
    End Class
End Namespace
