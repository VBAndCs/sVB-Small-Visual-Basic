Imports System.Reflection
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Library

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
                Case TokenType.Concatenation
                    methodInfo = scope.TypeInfoBag.Concat
                Case TokenType.Addition
                    methodInfo = scope.TypeInfoBag.Add
                Case TokenType.Subtraction
                    methodInfo = scope.TypeInfoBag.Subtract
                Case TokenType.Multiplication
                    methodInfo = scope.TypeInfoBag.Multiply
                Case TokenType.Division
                    methodInfo = scope.TypeInfoBag.Divide
                Case TokenType.Mod
                    methodInfo = scope.TypeInfoBag.Remainder
                Case TokenType.EqualsTo
                    methodInfo = scope.TypeInfoBag.EqualTo
                Case TokenType.NotEqualsTo
                    methodInfo = scope.TypeInfoBag.NotEqualTo
                Case TokenType.GreaterThan
                    methodInfo = scope.TypeInfoBag.GreaterThan
                Case TokenType.GreaterThanOrEqualsTo
                    methodInfo = scope.TypeInfoBag.GreaterThanOrEqualTo
                Case TokenType.LessThan
                    methodInfo = scope.TypeInfoBag.LessThan
                Case TokenType.LessThanOrEqualsTo
                    methodInfo = scope.TypeInfoBag.LessThanOrEqualTo
                Case TokenType.And, TokenType.Or
                    EmitShortCircuit([Operator].Type, scope)
                    Return
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

        Private Sub EmitShortCircuit(tokenType As TokenType, scope As CodeGenScope)
            Dim il = scope.ILGenerator

            If LeftHandSide Is Nothing Then
                LiteralExpression.Zero.EmitIL(scope)
            Else
                LeftHandSide.EmitIL(scope)
            End If

            il.Emit(OpCodes.Dup)
            il.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)

            Dim shortCircuitLabel = il.DefineLabel()

            If tokenType = TokenType.And Then
                il.Emit(OpCodes.Brfalse, shortCircuitLabel)
            Else
                il.Emit(OpCodes.Brtrue, shortCircuitLabel)
            End If

            il.Emit(OpCodes.Pop)

            If RightHandSide Is Nothing Then
                LiteralExpression.Zero.EmitIL(scope)
            Else
                RightHandSide.EmitIL(scope)
            End If

            il.MarkLabel(shortCircuitLabel)
        End Sub

        Public Overrides Function ToString() As String
            Return $"({LeftHandSide} {[Operator].Text} {RightHandSide})"
        End Function

        Public Overrides Function InferType(symbolTable As SymbolTable) As VariableType
            Dim leftType = LeftHandSide.InferType(symbolTable)

            If [Operator].IsIllegal OrElse RightHandSide Is Nothing Then
                Return leftType
            End If

            Dim rightType = RightHandSide.InferType(symbolTable)

            Select Case [Operator].Type
                Case TokenType.Division, TokenType.Mod
                    Return VariableType.Double

                Case TokenType.Multiplication ' We can use * to repeat a string
                    If leftType = VariableType.String OrElse rightType = VariableType.String OrElse
                            leftType = VariableType.Array OrElse rightType = VariableType.Array OrElse
                            leftType >= VariableType.Control OrElse rightType >= VariableType.Control Then
                        Return VariableType.String
                    Else
                        Return VariableType.Double
                    End If

                Case TokenType.Concatenation
                    Return VariableType.String

                Case TokenType.Addition
                    If leftType = VariableType.Any OrElse rightType = VariableType.Any OrElse
                            leftType = VariableType.String OrElse rightType = VariableType.String OrElse
                            leftType = VariableType.Array OrElse rightType = VariableType.Array OrElse
                            leftType >= VariableType.Control OrElse rightType >= VariableType.Control Then
                        Return VariableType.String
                    ElseIf leftType = VariableType.Date OrElse rightType = VariableType.Date Then
                        Return VariableType.Date
                    Else
                        Return VariableType.Double
                    End If

                Case TokenType.Subtraction
                    If leftType = VariableType.Date Then
                        Return VariableType.Date
                    Else
                        Return VariableType.Double
                    End If

                Case Else
                    Return VariableType.Boolean
            End Select
        End Function

        Public Overrides Function Evaluate(runner As Engine.ProgramRunner) As Primitive
            Dim leftExpr = LeftHandSide.Evaluate(runner)

            Select Case _Operator.Type
                Case TokenType.Or
                    If CBool(leftExpr) Then
                        Return New Primitive(True)
                    End If
                    Return RightHandSide.Evaluate(runner)

                Case TokenType.And
                    If CBool(leftExpr) Then
                        Return RightHandSide.Evaluate(runner)
                    End If
                    Return New Primitive(False)
            End Select

            Dim rightExpr = RightHandSide.Evaluate(runner)
            Select Case _Operator.Type
                Case TokenType.Concatenation
                    Return leftExpr.Concat(rightExpr)
                Case TokenType.Addition
                    Return leftExpr.Add(rightExpr)
                Case TokenType.Division
                    Return leftExpr.Divide(rightExpr)
                Case TokenType.Mod
                    Return leftExpr.Remainder(rightExpr)
                Case TokenType.EqualsTo
                    Return leftExpr.EqualTo(rightExpr)
                Case TokenType.GreaterThan
                    Return leftExpr.GreaterThan(rightExpr)
                Case TokenType.GreaterThanOrEqualsTo
                    Return leftExpr.GreaterThanOrEqualTo(rightExpr)
                Case TokenType.LessThan
                    Return leftExpr.LessThan(rightExpr)
                Case TokenType.LessThanOrEqualsTo
                    Return leftExpr.LessThanOrEqualTo(rightExpr)
                Case TokenType.Multiplication
                    Return leftExpr.Multiply(rightExpr)
                Case TokenType.NotEqualsTo
                    Return leftExpr.NotEqualTo(rightExpr)
                Case TokenType.Subtraction
                    Return leftExpr.Subtract(rightExpr)
                Case Else
                    Return Nothing
            End Select
        End Function

        Public Overrides Function ToVB() As String
            Return $"({LeftHandSide.ToVB()} {[Operator].Text} {RightHandSide.ToVB()})"
        End Function
    End Class
End Namespace
