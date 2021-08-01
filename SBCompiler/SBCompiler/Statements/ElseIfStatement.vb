Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Statements
    Public Class ElseIfStatement
        Inherits Statement

        Public ThenStatements As New List(Of Statement)()
        Public Condition As Expression
        Public ElseIfToken As TokenInfo
        Public ThenToken As TokenInfo

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            ElseIfToken.Parent = Me
            ThenToken.Parent = Me

            If Condition IsNot Nothing Then
                Condition.Parent = Me
                Condition.AddSymbols(symbolTable)
            End If

            For Each thenStatement In ThenStatements
                thenStatement.Parent = Me
                thenStatement.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub PrepareForEmit(ByVal scope As CodeGenScope)
            For Each thenStatement In ThenStatements
                thenStatement.PrepareForEmit(scope)
            Next
        End Sub

        Public Overrides Function GetIndentationLevel(ByVal line As Integer) As Integer
            If line = ElseIfToken.Line Then
                Return 0
            End If

            Dim statementContaining = GetStatementContaining(ThenStatements, line)

            If statementContaining IsNot Nothing Then
                Return 1 + statementContaining.GetIndentationLevel(line)
            End If

            Return 1
        End Function

        Public Overrides Function ToString() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0} {1}" & VisualBasic.Constants.vbCrLf, New Object(1) {ElseIfToken.Text, Condition})

            For Each thenStatement In ThenStatements
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "  {0}", New Object(0) {thenStatement})
            Next

            Return stringBuilder.ToString()
        End Function
    End Class
End Namespace
