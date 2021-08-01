Imports System.Globalization
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Statements
    Public Class ReturnStatement
        Inherits Statement

        Public ReturnExpression As Expression
        Friend Subroutine As SubroutineStatement

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            If ReturnExpression IsNot Nothing Then
                ReturnExpression.Parent = Me
            End If
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            Dim code = ""

            If Subroutine.SubToken.Token = Token.Function Then
                code = $"Stack.PushValue(""_sVB_ReturnValues"", {If(ReturnExpression, ChrW(34) & ChrW(34))})" & vbCrLf
            End If

            code &= $"GoTo _EXIT_SUB_{subroutine.Name.NormalizedText}"

            CodeGenerator.LowerAndEmit(code, scope, Subroutine)
        End Sub

        Public Overrides Function ToString() As String
            Return $"{StartToken} {ReturnExpression}"
        End Function
    End Class
End Namespace
