Imports System.Globalization
Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic.Statements
    Public Class SubroutineCallStatement
        Inherits Statement

        Public SubroutineName As TokenInfo

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim methodInfo = scope.MethodBuilders(SubroutineName.NormalizedText)
            scope.ILGenerator.EmitCall(OpCodes.Call, methodInfo, Nothing)
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0}()", New Object(0) {SubroutineName.Text})
        End Function
    End Class
End Namespace
