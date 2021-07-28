Imports System
Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public Class PropertyExpression
        Inherits Expression

        Public Property TypeName As TokenInfo
        Public Property PropertyName As TokenInfo

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            _TypeName.Parent = Me.Parent
            _PropertyName.Parent = Me.Parent
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim typeInfo = scope.TypeInfoBag.Types(TypeName.NormalizedText)
            Dim propertyInfo = typeInfo.Properties(PropertyName.NormalizedText)
            Dim getMethod As MethodInfo = propertyInfo.GetGetMethod()
            scope.ILGenerator.EmitCall(OpCodes.Call, getMethod, Nothing)
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0}.{1}", New Object(1) {TypeName.Text, PropertyName.Text})
        End Function
    End Class
End Namespace
