Imports System.Collections.Generic

Namespace Microsoft.SmallVisualBasic.Completion
    Public Class CompletionBag
        Public IsFirstToken As Boolean
        Public ShowCompletion As Boolean = True
        Public SubroutineName As String
        Public NextToEquals As Boolean
        Public SelectEspecialItem As String
        Friend NextToOperator As Boolean
        Public ForHelp As Boolean
        Public IsMethod As Boolean
        Public ReadOnly Property ParseTree As List(Of Statements.Statement)
        Public ReadOnly Property SymbolTable As SymbolTable
        Public ReadOnly Property TypeInfoBag As TypeInfoBag
        Public ReadOnly Property CompletionItems As List(Of CompletionItem)

        Public Sub New(
                      typeInfoBag As TypeInfoBag,
                      symbolTable As SymbolTable,
                      parseTree As List(Of Statements.Statement)
                 )

            _TypeInfoBag = typeInfoBag
            _SymbolTable = symbolTable
            _CompletionItems = New List(Of CompletionItem)()
            _ParseTree = parseTree
        End Sub
    End Class
End Namespace
