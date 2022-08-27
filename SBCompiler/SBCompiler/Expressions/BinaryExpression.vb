Imports System
Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic.Expressions
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

            LeftHandSide.EmitIL(scope)
            RightHandSide.EmitIL(scope)
            scope.ILGenerator.EmitCall(OpCodes.Call, methodInfo, Nothing)
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "({0} {1} {2})", New Object(2) {LeftHandSide, [Operator].Text, RightHandSide})
        End Function
    End Class
End Namespace
