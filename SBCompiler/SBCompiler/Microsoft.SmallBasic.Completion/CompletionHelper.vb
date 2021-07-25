Imports System.IO
Imports System.Reflection
Imports Microsoft.SmallBasic.Library
Imports Microsoft.SmallBasic.Statements

Namespace Microsoft.SmallBasic.Completion
    Public Class CompletionHelper
        Private _compiler As Compiler

        Public Sub New()
            _compiler = New Compiler()
        End Sub

        Public Function GetEmptyCompletionBag() As CompletionBag
            Return New CompletionBag(_compiler.TypeInfoBag, _compiler.Parser.SymbolTable)
        End Function

        Public Function GetCompletionItems(ByVal source As TextReader, ByVal line As Integer, ByVal column As Integer) As CompletionBag
            _compiler.Compile(source)
            Dim emptyCompletionBag As CompletionBag = GetEmptyCompletionBag()
            Dim statement As Statement = Nothing

            For num = _compiler.Parser.ParseTree.Count - 1 To 0 Step -1
                Dim statement2 = _compiler.Parser.ParseTree(num)

                If line >= statement2.StartToken.Line Then
                    statement = statement2
                    Exit For
                End If
            Next

            If statement IsNot Nothing Then
                FillCompletionItemsFromStatement(statement, emptyCompletionBag, line, column)
            Else
                FillAllGlobalItems(emptyCompletionBag, inGlobalScope:=True)
            End If

            emptyCompletionBag.CompletionItems.Sort(Function(ByVal ci1, ByVal ci2) ci1.DisplayName.CompareTo(ci2.DisplayName))
            Return emptyCompletionBag
        End Function

        Public Function GetIndentationLevel(ByVal source As TextReader, ByVal lineNumber As Integer) As Integer
            _compiler.Compile(source)
            Return GetIndentationLevel(lineNumber)
        End Function

        Public Function GetIndentationLevel(ByVal lineNumber As Integer) As Integer
            Return If(Statement.GetStatementContaining(_compiler.Parser.ParseTree, lineNumber)?.GetIndentationLevel(lineNumber), 0)
        End Function

        Private Sub FillCompletionItemsFromStatement(ByVal statement As Statement, ByVal completionBag As CompletionBag, ByVal line As Integer, ByVal column As Integer)
            statement.PopulateCompletionItems(completionBag, line, column, globalScope:=True)
        End Sub

        Public Shared Sub FillAllGlobalItems(ByVal completionBag As CompletionBag, ByVal inGlobalScope As Boolean)
            FillExpressionItems(completionBag)
            FillAllKeywords(completionBag)
            FillSubroutines(completionBag)

            If inGlobalScope Then
                FillKeywords(completionBag, Token.Sub, Token.Function)

            End If
        End Sub

        Public Shared Sub FillAllKeywords(ByVal completionBag As CompletionBag)
            FillKeywords(completionBag, Token.If, Token.For, Token.Goto, Token.While, Token.Return, Token.ExitLoop)
        End Sub

        Public Shared Sub FillLogicalExpressionItems(ByVal completionBag As CompletionBag)
            FillExpressionItems(completionBag)
            FillKeywords(completionBag, Token.And, Token.Or)
        End Sub

        Public Shared Sub FillExpressionItems(ByVal completionBag As CompletionBag)
            FillTypeNames(completionBag)
            FillVariables(completionBag)
        End Sub

        Public Shared Sub FillSubroutines(ByVal completionBag As CompletionBag)
            For Each subroutine In completionBag.SymbolTable.Subroutines
                Dim completionItem As CompletionItem = New CompletionItem()
                completionItem.Name = subroutine.Key
                completionItem.DisplayName = subroutine.Value.Text
                completionItem.ReplacementText = subroutine.Value.Text & "()"
                completionItem.ItemType = CompletionItemType.SubroutineName
                Dim item = completionItem
                completionBag.CompletionItems.Add(item)
            Next
        End Sub

        Public Shared Sub FillLabels(ByVal completionBag As CompletionBag)
            For Each label In completionBag.SymbolTable.Labels
                Dim completionItem As CompletionItem = New CompletionItem()
                completionItem.Name = label.Key
                completionItem.DisplayName = label.Value.Text
                completionItem.ItemType = CompletionItemType.Label
                Dim item = completionItem
                completionBag.CompletionItems.Add(item)
            Next
        End Sub

        Public Shared Sub FillVariables(ByVal completionBag As CompletionBag)
            For Each variable In completionBag.SymbolTable.Variables
                Dim completionItem As CompletionItem = New CompletionItem()
                completionItem.Name = variable.Key
                completionItem.DisplayName = variable.Value.Text
                completionItem.ItemType = CompletionItemType.Variable
                Dim item = completionItem
                completionBag.CompletionItems.Add(item)
            Next
        End Sub

        Public Shared Sub FillTypeNames(ByVal completionBag As CompletionBag)
            For Each type In completionBag.TypeInfoBag.Types

                If Not type.Value.HideFromIntellisense Then
                    Dim completionItem As CompletionItem = New CompletionItem()
                    completionItem.Name = type.Key
                    completionItem.DisplayName = type.Value.Type.Name
                    completionItem.ItemType = CompletionItemType.TypeName
                    completionItem.MemberInfo = type.Value.Type
                    Dim item = completionItem
                    completionBag.CompletionItems.Add(item)
                End If
            Next
        End Sub

        Public Shared Sub FillMemberNames(ByVal completionBag As CompletionBag, ByVal typeInfo As TypeInfo)
            For Each method In typeInfo.Methods

                If Not IsHiddenFromIntellisense(method.Value) Then
                    Dim completionItem As CompletionItem = New CompletionItem()
                    completionItem.Name = method.Key
                    completionItem.DisplayName = method.Value.Name
                    completionItem.ItemType = CompletionItemType.MethodName
                    completionItem.MemberInfo = method.Value
                    completionItem.ReplacementText = method.Value.Name & "("
                    If method.Value.GetParameters().Length = 0 Then
                        completionItem.ReplacementText &= ")"
                    End If

                    completionBag.CompletionItems.Add(completionItem)
                End If
            Next

            For Each [property] In typeInfo.Properties
                If Not IsHiddenFromIntellisense([property].Value) Then
                    Dim completionItem3 As CompletionItem = New CompletionItem()
                    completionItem3.Name = [property].Key
                    completionItem3.DisplayName = [property].Value.Name
                    completionItem3.ItemType = CompletionItemType.PropertyName
                    completionItem3.MemberInfo = [property].Value
                    Dim item = completionItem3
                    completionBag.CompletionItems.Add(item)
                End If
            Next

            For Each [event] In typeInfo.Events
                If Not IsHiddenFromIntellisense([event].Value) Then
                    Dim completionItem4 As New CompletionItem()
                    completionItem4.Name = [event].Key
                    completionItem4.DisplayName = [event].Value.Name
                    completionItem4.ItemType = CompletionItemType.EventName
                    completionItem4.MemberInfo = [event].Value
                    Dim item2 = completionItem4
                    completionBag.CompletionItems.Add(item2)
                End If
            Next
        End Sub

        Public Shared Sub FillKeywords(ByVal completionBag As CompletionBag, ParamArray keywords As Token())
            For Each token In keywords
                Dim completionItem As CompletionItem = New CompletionItem()
                completionItem.Name = token.ToString()
                completionItem.DisplayName = token.ToString()
                completionItem.ItemType = CompletionItemType.Keyword
                Dim item = completionItem
                completionBag.CompletionItems.Add(item)
            Next
        End Sub

        Private Shared Function IsHiddenFromIntellisense(ByVal mi As MemberInfo) As Boolean
            Return mi.GetCustomAttributes(GetType(HideFromIntellisenseAttribute), inherit:=False).Length > 0
        End Function
    End Class
End Namespace
