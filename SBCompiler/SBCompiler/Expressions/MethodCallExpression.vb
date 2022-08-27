Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Text

Namespace Microsoft.SmallBasic.Expressions
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
        Public Property OuterSubRoutine As Statements.SubroutineStatement

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
            symbolTable.AllIdentifiers.Add(_TypeName)

            _MethodName.SymbolType = If(_TypeName.IsIllegal, Completion.CompletionItemType.SubroutineName, Completion.CompletionItemType.MethodName)
            symbolTable.AllIdentifiers.Add(_MethodName)

            symbolTable.FixNames(_TypeName, _MethodName, True)
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If TypeName.Type = TokenType.Illegal Then ' Function Call
                Dim subroutine As New Statements.SubroutineCallStatement() With {
                    .Name = MethodName,
                    .Args = Arguments,
                    .IsFunctionCall = True,
                    .OuterSubroutine = OuterSubRoutine
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
            Dim typeInfo = scope.TypeInfoBag.Types(TypeName.NormalizedText)
            Return typeInfo.Methods(MethodName.NormalizedText)
        End Function

        Public Overrides Function ToString() As String
            Dim stringBuilder As New StringBuilder()
            If TypeName.Type = TokenType.Illegal Then
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{1}(", New Object(1) {TypeName.Text, MethodName.Text})
            Else
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}.{1}(", New Object(1) {TypeName.Text, MethodName.Text})
            End If

            If Arguments.Count > 0 Then
                For Each argument In Arguments
                    stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}, ", New Object(0) {argument})
                Next

                stringBuilder.Remove(stringBuilder.Length - 2, 2)
            End If

            stringBuilder.Append(")")
            Return stringBuilder.ToString()
        End Function
    End Class
End Namespace
