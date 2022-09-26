Imports System
Imports System.Globalization
Imports System.Reflection.Emit

Namespace Microsoft.SmallVisualBasic.Expressions
    <Serializable>
    Public Class LiteralExpression
        Inherits Expression

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(literal As Token)
            _Literal = literal
        End Sub

        Public Property Literal As Token

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            _Literal.Parent = Me.Parent
            If Literal.Type = TokenType.DateLiteral Then
                Dim result = Parser.ParseDateLiteral(Literal.Text)
                If result.Ticks = 0 Then
                    Dim kind = If(result.IsDate, "date", "time span")
                    symbolTable.Errors.Add(New [Error](Literal, $"Invalid {kind} format!"))
                End If
            End If
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            Select Case Literal.Type
                Case TokenType.StringLiteral
                    scope.ILGenerator.Emit(OpCodes.Ldstr, GetString(Literal.Text, """"))
                    scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.StringToPrimitive, Nothing)

                Case TokenType.DateLiteral
                    Dim result = Parser.ParseDateLiteral(Literal.Text)
                    scope.ILGenerator.Emit(OpCodes.Ldc_I8, result.Ticks)
                    scope.ILGenerator.EmitCall(OpCodes.Call, If(result.IsDate, scope.TypeInfoBag.DateToPrimitive, scope.TypeInfoBag.TimeSpanToPrimitive), Nothing)

                Case TokenType.True
                    scope.ILGenerator.Emit(OpCodes.Ldstr, "True")
                    scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.StringToPrimitive, Nothing)

                Case TokenType.False
                    scope.ILGenerator.Emit(OpCodes.Ldstr, "False")
                    scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.StringToPrimitive, Nothing)

                Case TokenType.NumericLiteral
                    scope.ILGenerator.Emit(OpCodes.Ldc_R8, Double.Parse(Literal.Text, CultureInfo.InvariantCulture))
                    scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.NumberToPrimitive, Nothing)
            End Select
        End Sub


        Private Function GetString(literal As String, enclosingChar As Char) As String
            Dim sb As New Text.StringBuilder()
            For i = 1 To literal.Length - 2
                If literal(i) = enclosingChar Then i += 1 ' escap
                sb.Append(literal(i))
            Next
            Return sb.ToString
        End Function

        Public Overrides Function ToString() As String
            Return Literal.Text
        End Function
    End Class

End Namespace
