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

        Private _arguments As List(Of Expression) = New List(Of Expression)()
        Public Property TypeName As TokenInfo
        Public Property MethodName As TokenInfo

        Public ReadOnly Property Arguments As List(Of Expression)
            Get
                Return _arguments
            End Get
        End Property

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            For Each argument In Arguments
                argument.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim methodInfo = GetMethodInfo(scope)

            For Each argument In Arguments
                argument.EmitIL(scope)
            Next

            scope.ILGenerator.EmitCall(OpCodes.Call, methodInfo, Nothing)
        End Sub

        Public Function GetMethodInfo(ByVal scope As CodeGenScope) As MethodInfo
            Dim typeInfo = scope.TypeInfoBag.Types(TypeName.NormalizedText)
            Return typeInfo.Methods(MethodName.NormalizedText)
        End Function

        Public Overrides Function ToString() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}.{1}(", New Object(1) {TypeName.Text, MethodName.Text})

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
