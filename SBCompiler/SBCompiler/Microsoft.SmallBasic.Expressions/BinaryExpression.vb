Imports System
Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public Class BinaryExpression
        Inherits Expression

        Public Property LeftHandSide As Expression
        Public Property [Operator] As TokenInfo
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

            Select Case [Operator].Token
                Case Token.Addition
                    methodInfo = scope.TypeInfoBag.Add
                Case Token.Subtraction
                    methodInfo = scope.TypeInfoBag.Subtract
                Case Token.Multiplication
                    methodInfo = scope.TypeInfoBag.Multiply
                Case Token.Division
                    methodInfo = scope.TypeInfoBag.Divide
                Case Token.Equals
                    methodInfo = scope.TypeInfoBag.EqualTo
                Case Token.NotEqualTo
                    methodInfo = scope.TypeInfoBag.NotEqualTo
                Case Token.GreaterThan
                    methodInfo = scope.TypeInfoBag.GreaterThan
                Case Token.GreaterThanEqualTo
                    methodInfo = scope.TypeInfoBag.GreaterThanOrEqualTo
                Case Token.LessThan
                    methodInfo = scope.TypeInfoBag.LessThan
                Case Token.LessThanEqualTo
                    methodInfo = scope.TypeInfoBag.LessThanOrEqualTo
                Case Token.And
                    methodInfo = scope.TypeInfoBag.And
                Case Token.Or
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
