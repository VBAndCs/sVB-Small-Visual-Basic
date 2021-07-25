Imports System.Globalization
Imports System.Reflection.Emit
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Statements
    Public Class SubroutineCallStatement
        Inherits Statement

        Public Name As TokenInfo
        Public Args As List(Of Expression)
        Public IsFunctionCall As Boolean

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            For Each arg In Args
                arg.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If Args.Count > 0 Then
                Dim code As New Text.StringBuilder()
                For i = Args.Count - 1 To 0 Step -1
                    code.AppendLine($"Stack.PushValue(""_sVB_Arguments"", {Args(i)})")
                Next

                LowerAndEmit(code.ToString(), scope)
            End If

            Dim methodInfo = scope.MethodBuilders(Name.NormalizedText)
            scope.ILGenerator.EmitCall(OpCodes.Call, methodInfo, Nothing)

            If IsFunctionCall Then
                MethodCallStatement.DoNotPopReturnValue = True
                LowerAndEmit($"Stack.PopValue(""_sVB_ReturnValues"")", scope)
                MethodCallStatement.DoNotPopReturnValue = False
            End If
        End Sub

        Private Sub LowerAndEmit(code As String, scope As CodeGenScope)

            Dim _parser = Parser.Parse(code)

            ' EmitIL
            For Each item In _parser.ParseTree
                item.PrepareForEmit(scope)
            Next

            For Each item In _parser.ParseTree
                item.EmitIL(scope)
            Next
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0}()", New Object(0) {Name.Text})
        End Function
    End Class
End Namespace
