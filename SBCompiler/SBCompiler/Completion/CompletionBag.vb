Imports System.Collections.Generic

Namespace Microsoft.SmallBasic.Completion
    Public Class CompletionBag
        Private _typeInfoBag As TypeInfoBag
        Private _symbolTable As SymbolTable
        Private _completionItems As List(Of CompletionItem)

        Public ReadOnly Property SymbolTable As SymbolTable
            Get
                Return _symbolTable
            End Get
        End Property

        Public ReadOnly Property TypeInfoBag As TypeInfoBag
            Get
                Return _typeInfoBag
            End Get
        End Property

        Public ReadOnly Property CompletionItems As List(Of CompletionItem)
            Get
                Return _completionItems
            End Get
        End Property

        Public Sub New(
                      typeInfoBag As TypeInfoBag,
                      symbolTable As SymbolTable
                 )
            _typeInfoBag = typeInfoBag
            _symbolTable = symbolTable
            _completionItems = New List(Of CompletionItem)()
        End Sub
    End Class
End Namespace
