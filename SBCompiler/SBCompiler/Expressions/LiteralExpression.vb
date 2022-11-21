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

        Public Shared ReadOnly Zero As New LiteralExpression(
            New Token() With {
                    .Type = TokenType.NumericLiteral,
                    .Text = 0
            }
        )

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            _Literal.Parent = Me.Parent

            Select Case Literal.Type
                Case TokenType.StringLiteral
                    If Not Literal.Text.EndsWith("""") Then
                        symbolTable.Errors.Add(New [Error](Literal, "A string literal must end with a quote."))
                    End If

                Case TokenType.DateLiteral
                    If Not Literal.Text.EndsWith("#") Then
                        symbolTable.Errors.Add(New [Error](Literal, "Date and duration literals must end with #."))
                    End If

                    Dim result = Parser.ParseDateLiteral(Literal.Text)
                    If Not result.Ticks.HasValue Then
                        Dim kind = If(result.IsDate, "date", "time span")
                        symbolTable.Errors.Add(New [Error](Literal, $"Invalid {kind} format!"))
                    End If
            End Select
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp Then
                ' no need to emit the exact literal. This is just for global file help,
                ' and we know what we need about the array from the symbol table
                scope.ILGenerator.Emit(OpCodes.Ldc_R8, 0.0)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.NumberToPrimitive, Nothing)
                Return
            End If

            Select Case Literal.Type
                Case TokenType.StringLiteral
                    scope.ILGenerator.Emit(OpCodes.Ldstr, GetString(Literal.Text, """"))
                    scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.StringToPrimitive, Nothing)

                Case TokenType.DateLiteral
                    Dim result = Parser.ParseDateLiteral(Literal.Text)
                    scope.ILGenerator.Emit(OpCodes.Ldc_I8, result.Ticks.Value)
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

        Public Overrides Function InferType(symbolTable As SymbolTable) As VariableType
            Select Case Literal.Type
                Case TokenType.StringLiteral
                    Return VariableType.String

                Case TokenType.DateLiteral
                    Return VariableType.Date

                Case TokenType.True
                    Return VariableType.Boolean

                Case TokenType.False
                    Return VariableType.Boolean

                Case TokenType.NumericLiteral
                    Return VariableType.Double

                Case Else
                    Return VariableType.Any
            End Select

        End Function
    End Class

End Namespace
