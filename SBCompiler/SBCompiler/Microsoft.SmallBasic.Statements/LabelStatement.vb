Imports System.Globalization
Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic.Statements
    Public Class LabelStatement
        Inherits Statement

        Public LabelToken As TokenInfo
        Public ColonToken As TokenInfo
        Public subroutine As SubroutineStatement

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            LabelToken.Parent = Me
            ColonToken.Parent = Me

            If LabelToken.Token <> 0 Then
                symbolTable.AddLabelDefinition(LabelToken)
            End If
        End Sub

        Public Overrides Sub PrepareForEmit(ByVal scope As CodeGenScope)
            Dim value As Label = scope.ILGenerator.DefineLabel()
            scope.Labels.Add(LabelToken.NormalizedText, value)
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim loc = scope.Labels(LabelToken.NormalizedText)
            scope.ILGenerator.MarkLabel(loc)
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0}{1}" & VisualBasic.Constants.vbCrLf, New Object(1) {LabelToken.Text, ColonToken.Text})
        End Function
    End Class
End Namespace
