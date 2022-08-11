Imports System.Collections.Generic
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.Statements
    Public MustInherit Class Statement
        Public Property StartToken As TokenInfo
        Public Property EndingComment As TokenInfo

        Public Property Parent As Statement

        Public Overridable Sub AddSymbols(symbolTable As SymbolTable)
            _StartToken.Parent = Me
            _EndingComment.Parent = Me
        End Sub

        Public MustOverride Function GetStatement(lineNumber) As Statement

        Public Overridable Sub PrepareForEmit(ByVal scope As CodeGenScope)
        End Sub

        Public Overridable Sub EmitIL(ByVal scope As CodeGenScope)
        End Sub

        Public Overridable Sub PopulateCompletionItems(completionBag As CompletionBag, ByVal line As Integer, ByVal column As Integer, ByVal globalScope As Boolean)
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

        Public Overridable Function GetKeywords() As LegalTokens
            Return Nothing
        End Function

        Public Function IsOfType(statementType As Type) As Boolean
            If statementType Is Nothing Then
                Select Case Me.GetType
                    Case GetType(SubroutineStatement),
                             GetType(IfStatement),
                             GetType(ForStatement),
                             GetType(WhileStatement)
                        Return True
                    Case Else
                        Return False
                End Select
            End If
            Return Me.GetType Is statementType
        End Function
    End Class
End Namespace
