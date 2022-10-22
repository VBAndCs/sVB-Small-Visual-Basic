Imports System.Collections.Generic
Imports Microsoft.SmallVisualBasic.Statements

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
        Friend IsHandler As Boolean
        Public CtrlSpace As Boolean
        Public ReadOnly Property ParseTree As List(Of Statement)
        Public ReadOnly Property GlobalParseTree As List(Of Statement)
        Public ReadOnly Property SymbolTable As SymbolTable
        Public ReadOnly Property GlobalSymbolTable As SymbolTable
        Public ReadOnly Property TypeInfoBag As TypeInfoBag
        Public ReadOnly Property CompletionItems As List(Of CompletionItem)

        Public Sub New(
                      typeInfoBag As TypeInfoBag,
                      symbolTable As SymbolTable,
                      globalSymbolTable As SymbolTable,
                      parseTree As List(Of Statement),
                     globalParseTree As List(Of Statement)
                 )

            _TypeInfoBag = typeInfoBag
            _SymbolTable = symbolTable
            _GlobalSymbolTable = globalSymbolTable
            _CompletionItems = New List(Of CompletionItem)()
            _ParseTree = parseTree
            _GlobalParseTree = globalParseTree
        End Sub
    End Class
End Namespace
