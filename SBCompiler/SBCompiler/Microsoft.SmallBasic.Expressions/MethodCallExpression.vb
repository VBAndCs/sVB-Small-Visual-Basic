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

        Public Sub New(startToken As TokenInfo, precedence As Integer, typeName As TokenInfo, methodName As TokenInfo, arguments As List(Of Expression))
            Me.StartToken = startToken
            Me.Precedence = precedence
            _TypeName = typeName
            _MethodName = methodName
            _Arguments = arguments
        End Sub

        Public Property TypeName As TokenInfo
        Public Property MethodName As TokenInfo
        Public Property OuterSubRoutine As Statements.SubroutineStatement


        Public ReadOnly Property Arguments As New List(Of Expression)

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            For Each argument In Arguments
                argument.Parent = Me.Parent
                argument.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            If TypeName.Token = Token.Illegal Then ' Function Call
                Dim subroutine As New Statements.SubroutineCallStatement() With {
                    .Name = MethodName,
                    .Args = Arguments,
                    .IsFunctionCall = True,
                    .OuterSubRoutine = OuterSubRoutine
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

        Public Function GetMethodInfo(ByVal scope As CodeGenScope) As MethodInfo
            Dim typeInfo = scope.TypeInfoBag.Types(TypeName.NormalizedText)
            Return typeInfo.Methods(MethodName.NormalizedText)
        End Function

        Public Overrides Function ToString() As String
            Dim stringBuilder As New StringBuilder()
            If TypeName.Token = Token.Illegal Then
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
