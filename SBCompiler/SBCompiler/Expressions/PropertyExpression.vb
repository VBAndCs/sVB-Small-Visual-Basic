Imports System
Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit
Imports Microsoft.SmallBasic.Statements

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public Class PropertyExpression
        Inherits Expression

        Public Property TypeName As TokenInfo
        Public Property PropertyName As TokenInfo

        Public IsDynamic As Boolean
        Public isSet As Boolean

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            _TypeName.Parent = Me.Parent
            _PropertyName.Parent = Me.Parent
            Dim name = TypeName.NormalizedText
            If IsDynamic OrElse name.StartsWith("data") Or name.EndsWith("data") Then
                symbolTable.AddDynamic(Me)
            End If
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            If IsDynamic Then
                Dim code = $"_sVB_dynamic_Data_ = {TypeName.Text}[""{PropertyName.Text}""]"
                Dim subroutine = SubroutineStatement.GetSubroutine(Me)
                If subroutine Is Nothing Then subroutine = SubroutineStatement.Current
                ArrayExpression.ParseAndEmit(code, subroutine, scope).EmitIL(scope)
            Else
                Dim typeInfo = scope.TypeInfoBag.Types(TypeName.NormalizedText)
                Dim propertyInfo = typeInfo.Properties(PropertyName.NormalizedText)
                Dim getMethod = propertyInfo.GetGetMethod()
                scope.ILGenerator.EmitCall(OpCodes.Call, getMethod, Nothing)
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return $"{TypeName.Text}{If(IsDynamic, "!", ".")}{PropertyName.Text}"
        End Function
    End Class
End Namespace
