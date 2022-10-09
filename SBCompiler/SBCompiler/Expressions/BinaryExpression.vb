Imports System.Reflection
Imports System.Reflection.Emit

Namespace Microsoft.SmallVisualBasic.Expressions
    <Serializable>
    Public Class BinaryExpression
        Inherits Expression

        Public Property LeftHandSide As Expression
        Public Property [Operator] As Token
        Public Property RightHandSide As Expression

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            _Operator.Parent = Me.Parent

            If LeftHandSide IsNot Nothing Then
                LeftHandSide.Parent = Me.Parent
                LeftHandSide.AddSymbols(symbolTable)
            End If

            If RightHandSide IsNot Nothing Then
                RightHandSide.Parent = Me.Parent
                RightHandSide.AddSymbols(symbolTable)
            End If
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)

            Dim methodInfo As MethodInfo = Nothing

            Select Case [Operator].Type
                Case TokenType.Addition
                    methodInfo = scope.TypeInfoBag.Add
                Case TokenType.Subtraction
                    methodInfo = scope.TypeInfoBag.Subtract
                Case TokenType.Multiplication
                    methodInfo = scope.TypeInfoBag.Multiply
                Case TokenType.Division
                    methodInfo = scope.TypeInfoBag.Divide
                Case TokenType.Equals
                    methodInfo = scope.TypeInfoBag.EqualTo
                Case TokenType.NotEqualTo
                    methodInfo = scope.TypeInfoBag.NotEqualTo
                Case TokenType.GreaterThan
                    methodInfo = scope.TypeInfoBag.GreaterThan
                Case TokenType.GreaterThanEqualTo
                    methodInfo = scope.TypeInfoBag.GreaterThanOrEqualTo
                Case TokenType.LessThan
                    methodInfo = scope.TypeInfoBag.LessThan
                Case TokenType.LessThanEqualTo
                    methodInfo = scope.TypeInfoBag.LessThanOrEqualTo
                Case TokenType.And
                    methodInfo = scope.TypeInfoBag.And
                Case TokenType.Or
                    methodInfo = scope.TypeInfoBag.Or
            End Select

            If LeftHandSide Is Nothing Then
                ' This is just for intellisense info of the global file
                ' that may contain errors which is igmored at design time.
                ' This will not affect the compilation.
                LiteralExpression.Zero.EmitIL(scope)
            Else
                LeftHandSide.EmitIL(scope)
            End If

            If RightHandSide Is Nothing Then
                LiteralExpression.Zero.EmitIL(scope)
            Else
                RightHandSide.EmitIL(scope)
            End If

            scope.ILGenerator.EmitCall(OpCodes.Call, methodInfo, Nothing)
        End Sub

        Public Overrides Function ToString() As String
            Return $"({LeftHandSide} {[Operator].Text} {RightHandSide})"
        End Function

        Public Overrides Function InferType(symbolTable As SymbolTable) As VariableType
            Dim leftType = LeftHandSide.InferType(symbolTable)
            If leftType = VariableType.None OrElse leftType >= VariableType.Control Then
                Return VariableType.None
            End If

            If [Operator].IsIllegal OrElse RightHandSide Is Nothing Then Return leftType

            Dim rightType = RightHandSide.InferType(symbolTable)
            If rightType = VariableType.None OrElse rightType >= VariableType.Control Then
                Return VariableType.None
            End If

            Select Case [Operator].Type
                Case TokenType.Multiplication, TokenType.Division
                    Return VariableType.Double

                Case TokenType.Addition
                    If leftType = VariableType.String OrElse
                            rightType = VariableType.String OrElse
                            leftType = VariableType.Array OrElse
                            rightType = VariableType.Array Then
                        Return VariableType.String

                    ElseIf leftType = VariableType.Date OrElse rightType = VariableType.Date Then
                        Return VariableType.Date
                    Else
                        Return VariableType.Double
                    End If

                Case TokenType.Sub
                    If leftType = VariableType.Date Then
                        Return VariableType.Date
                    Else
                        Return VariableType.Double
                    End If

                Case Else
                    Return VariableType.Boolean
            End Select

        End Function
    End Class
End Namespace
