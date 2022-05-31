Imports System.Collections.Generic
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.Statements
    Public MustInherit Class Statement
        Public Property StartToken As TokenInfo
        Public Property EndingComment As TokenInfo

        Public Parent As Statement

        Public Overridable Sub AddSymbols(ByVal symbolTable As SymbolTable)
            _StartToken.Parent = Me
            _EndingComment.Parent = Me
        End Sub

        Public Overridable Sub PrepareForEmit(ByVal scope As CodeGenScope)
        End Sub

        Public Overridable Sub EmitIL(ByVal scope As CodeGenScope)
        End Sub

        Public Overridable Sub PopulateCompletionItems(ByVal completionBag As CompletionBag, ByVal line As Integer, ByVal column As Integer, ByVal globalScope As Boolean)
            CompletionHelper.FillAllGlobalItems(completionBag, globalScope)
        End Sub

        Public Shared Function GetStatementContaining(statements As List(Of Statement), line As Integer) As Statement
            For num = statements.Count - 1 To 0 Step -1
                Dim statement = statements(num)
                Dim L = statement.StartToken.Line
                If L > 0 AndAlso line >= L Then Return statement
            Next
            Return Nothing
        End Function
    End Class
End Namespace
