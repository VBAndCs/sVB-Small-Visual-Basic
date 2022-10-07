Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Text

Namespace Microsoft.SmallVisualBasic.Expressions
    <Serializable>
    Public Class MethodCallExpression
        Inherits Expression

        Public Sub New(startToken As Token, precedence As Integer, typeName As Token, methodName As Token, arguments As List(Of Expression))
            Me.StartToken = startToken
            Me.Precedence = precedence
            _TypeName = typeName
            _MethodName = methodName
            _Arguments = arguments
        End Sub

        Public Property TypeName As Token
        Public Property MethodName As Token
        Public Property OuterSubroutine As Statements.SubroutineStatement

        Public ReadOnly Property Arguments As New List(Of Expression)

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            _TypeName.Parent = Me.Parent
            _MethodName.Parent = Me.Parent

            For Each argument In Arguments
                argument.Parent = Me.Parent
                argument.AddSymbols(symbolTable)
            Next

            _TypeName.SymbolType = Completion.CompletionItemType.GlobalVariable
            symbolTable.AddIdentifier(_TypeName)

            _MethodName.SymbolType = If(_TypeName.IsIllegal, Completion.CompletionItemType.SubroutineName, Completion.CompletionItemType.MethodName)
            symbolTable.AddIdentifier(_MethodName)

            symbolTable.FixNames(_TypeName, _MethodName, True)
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp Then
                LiteralExpression.Zero.EmitIL(scope)
                Return
            End If

            If TypeName.Type = TokenType.Illegal Then ' Function Call
                Dim subroutine As New Statements.SubroutineCallStatement() With {
                    .Name = MethodName,
                    .Args = Arguments,
                    .IsFunctionCall = True,
                    .OuterSubroutine = OuterSubroutine,
                    .KeepReturnValue = True
                }
                subroutine.EmitIL(scope)

            Else
                Dim methodInfo = GetMethodInfo(scope)
                For Each argument In Arguments
                    argument.EmitIL(scope)
                Next
                scope.ILGenerator.EmitCall(OpCodes.Call, methodInfo, Nothing)
            End If
        End Sub

        Public Function GetMethodInfo(scope As CodeGenScope) As MethodInfo
            Dim typeInfo = scope.TypeInfoBag.Types(TypeName.LCaseText)
            Return typeInfo.Methods(MethodName.LCaseText)
        End Function

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder()

            If Not TypeName.IsIllegal Then
                sb.Append($"{TypeName.Text}.")
            End If

            sb.Append(MethodName.Text & "(")

            If Arguments.Count > 0 Then
                For Each arg In Arguments
                    sb.Append(arg.ToString())
                    sb.Append(", ")
                Next

                sb.Remove(sb.Length - 2, 2)
            End If

            sb.Append(")")
            Return sb.ToString()
        End Function
    End Class
End Namespace
