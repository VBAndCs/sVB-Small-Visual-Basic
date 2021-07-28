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

        Public Shared Sub FillAllGlobalItems(completionBag As CompletionBag, inGlobalScope As Boolean)
            FillExpressionItems(completionBag)
            FillAllKeywords(completionBag)
            FillSubroutines(completionBag)

            If inGlobalScope Then
                FillLocals(completionBag, "")
                FillKeywords(completionBag, Token.Sub, Token.Function)
            End If
        End Sub

        Public Shared Sub FillAllKeywords(ByVal completionBag As CompletionBag)
            FillKeywords(completionBag, Token.If, Token.For, Token.Goto, Token.While, Token.Return, Token.ExitLoop, Token.ContinueLoop)
        End Sub

        Public Shared Sub FillLogicalExpressionItems(ByVal completionBag As CompletionBag)
            FillExpressionItems(completionBag)
            FillKeywords(completionBag, Token.And, Token.Or)
        End Sub

        Public Shared Sub FillExpressionItems(bag As CompletionBag)
            FillTypeNames(bag)
            FillVariables(bag)
            FillBooleanLitrals(bag)
        End Sub

        Public Shared Sub FillSubroutines(bag As CompletionBag)
            For Each subroutine In bag.SymbolTable.Subroutines
                bag.CompletionItems.Add(New CompletionItem() With {
                    .Name = subroutine.Key,
                    .DisplayName = subroutine.Value.Text,
                    .ReplacementText = subroutine.Value.Text & "()",
                    .ItemType = CompletionItemType.SubroutineName
                })
            Next
        End Sub

        Public Shared Sub FillLabels(bag As CompletionBag)
            For Each label In bag.SymbolTable.Labels
                bag.CompletionItems.Add(New CompletionItem() With {
                    .Name = label.Key,
                    .DisplayName = label.Value.Text,
                    .ItemType = CompletionItemType.Label
               })
            Next
        End Sub

        Public Shared Sub FillBooleanLitrals(bag As CompletionBag)
            bag.CompletionItems.Add(New CompletionItem() With {
                .Name = "True",
                .DisplayName = "True",
                .ItemType = CompletionItemType.Keyword
            })

            bag.CompletionItems.Add(New CompletionItem() With {
                .Name = "False",
                .DisplayName = "False",
                .ItemType = CompletionItemType.Keyword
            })
        End Sub


        Public Shared Sub FillVariables(completionBag As CompletionBag)
            For Each variable In completionBag.SymbolTable.Variables
                completionBag.CompletionItems.Add(New CompletionItem() With {
                    .Name = variable.Key,
                    .DisplayName = variable.Value.Text,
                    .ItemType = CompletionItemType.Variable
                })
            Next
        End Sub

        Public Shared Sub FillLocals(completionBag As CompletionBag, subroutine As String)
            For Each variable In completionBag.SymbolTable.Locals
                Dim addToBag = False
                If subroutine = "" Then
                    If Not variable.Key.Contains(".") Then
                        addToBag = True
                    End If
                ElseIf variable.Key.StartsWith(subroutine & ".") Then
                    addToBag = True
                End If

                If addToBag Then
                    completionBag.CompletionItems.Add(New CompletionItem() With {
                        .Name = variable.Key,
                        .DisplayName = variable.Value.Identifier.Text,
                        .ItemType = CompletionItemType.Variable
                    })
                End If
            Next
        End Sub


        Public Shared Sub FillTypeNames(bag As CompletionBag)
            For Each type In bag.TypeInfoBag.Types
                If Not type.Value.HideFromIntellisense Then
                    bag.CompletionItems.Add(New CompletionItem() With {
                        .Name = type.Key,
                        .DisplayName = type.Value.Type.Name,
                        .ItemType = CompletionItemType.TypeName,
                        .MemberInfo = type.Value.Type
                    })
                End If
            Next
        End Sub

        Public Shared Sub FillMemberNames(ByVal completionBag As CompletionBag, ByVal typeInfo As TypeInfo)
            For Each method In typeInfo.Methods

                If Not IsHiddenFromIntellisense(method.Value) Then
                    Dim completionItem As New CompletionItem() With {
                        .Name = method.Key,
                        .DisplayName = method.Value.Name,
                        .ItemType = CompletionItemType.MethodName,
                        .MemberInfo = method.Value,
                        .ReplacementText = method.Value.Name & "("
                    }

                    If method.Value.GetParameters().Length = 0 Then
                        completionItem.ReplacementText &= ")"
                    End If

                    completionBag.CompletionItems.Add(completionItem)
                End If
            Next

            For Each [property] In typeInfo.Properties
                If Not IsHiddenFromIntellisense([property].Value) Then
                    completionBag.CompletionItems.Add(New CompletionItem() With {
                        .Name = [property].Key,
                        .DisplayName = [property].Value.Name,
                        .ItemType = CompletionItemType.PropertyName,
                        .MemberInfo = [property].Value
                    })
                End If
            Next

            For Each [event] In typeInfo.Events
                If Not IsHiddenFromIntellisense([event].Value) Then
                    completionBag.CompletionItems.Add(New CompletionItem() With {
                        .Name = [event].Key,
                        .DisplayName = [event].Value.Name,
                        .ItemType = CompletionItemType.EventName,
                        .MemberInfo = [event].Value
                    })
                End If
            Next
        End Sub

        Public Shared Sub FillWords(completionBag As CompletionBag, ParamArray words As String())
            For Each word In words
                completionBag.CompletionItems.Add(New CompletionItem() With {
                    .Name = word,
                    .DisplayName = word,
                    .ItemType = CompletionItemType.Keyword
                })
            Next
        End Sub

        Public Shared Sub FillKeywords(completionBag As CompletionBag, ParamArray keywords As Token())
            For Each token In keywords
                completionBag.CompletionItems.Add(New CompletionItem() With {
                    .Name = token.ToString(),
                    .DisplayName = token.ToString(),
                    .ItemType = CompletionItemType.Keyword
                })
            Next
        End Sub

        Private Shared Function IsHiddenFromIntellisense(ByVal mi As MemberInfo) As Boolean
            Return mi.GetCustomAttributes(GetType(HideFromIntellisenseAttribute), inherit:=False).Length > 0
        End Function
    End Class
End Namespace
