Imports System.Text
Imports Microsoft.SmallVisualBasic.Expressions

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class ElseIfStatement
        Inherits Statement

        Public ThenStatements As New List(Of Statement)()
        Public Condition As Expression
        Public ElseIfToken As Token
        Public ThenToken As Token

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber <= ThenToken.Line Then Return Me

            For Each statment In ThenStatements
                Dim st = statment.GetStatementAt(lineNumber)
                If st IsNot Nothing Then Return st
            Next

            Return Nothing
        End Function

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
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

        Public Overrides Sub PrepareForEmit(scope As CodeGenScope)
            If scope.ForGlobalHelp Then Return

            For Each thenStatement In ThenStatements
                thenStatement.PrepareForEmit(scope)
            Next
        End Sub

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder()
            sb.AppendLine($"{ElseIfToken.Text} {Condition} {ThenToken.Text}")

            For Each st In ThenStatements
                sb.Append("  ")
                sb.Append(st.ToString())
            Next

            Return sb.ToString()
        End Function

        Public Overrides Sub InferType(symbolTable As SymbolTable)
            For Each st In ThenStatements
                st.InferType(symbolTable)
            Next
        End Sub
    End Class
End Namespace
