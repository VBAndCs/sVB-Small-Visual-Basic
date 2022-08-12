Imports System.IO
Imports System.Reflection
Imports Microsoft.SmallBasic.Library
Imports Microsoft.SmallBasic.Statements

Namespace Microsoft.SmallBasic.Completion
    Public Class CompletionHelper
        Private _compiler As Compiler
        Public Shared DoNotAddGlobals As Boolean

        Public Sub New()
            _compiler = New Compiler()
        End Sub

        Public Function GetEmptyCompletionBag() As CompletionBag
            Return New CompletionBag(
                _compiler.TypeInfoBag,
                _compiler.Parser.SymbolTable
            )
        End Function

        Public Function GetCompletionItems(
                        source As TextReader,
                        line As Integer,
                        column As Integer
                    ) As CompletionBag


            _compiler.Compile(source, True)

            Dim completionBag = GetEmptyCompletionBag()
            Dim statement = GetStatement(_compiler, line)

            If statement IsNot Nothing Then
                FillCompletionItemsFromStatement(statement, completionBag, line, column)
            Else
                FillAllGlobalItems(completionBag, inGlobalScope:=True)
            End If

            completionBag.CompletionItems.Sort(Function(ByVal ci1, ByVal ci2) ci1.DisplayName.CompareTo(ci2.DisplayName))
            Return completionBag
        End Function

        Public Shared Function GetStatement(
                        compiler As Compiler,
                        line As Integer
                  ) As Statement

            For i = compiler.Parser.ParseTree?.Count - 1 To 0 Step -1
                Dim statement = compiler.Parser.ParseTree(i)

                If line >= statement.StartToken.Line Then
                    Return statement
                    Exit For
                End If
            Next

            Return Nothing
        End Function

        Private Sub FillCompletionItemsFromStatement(statement As Statement, completionBag As CompletionBag, line As Integer, column As Integer)
            statement.PopulateCompletionItems(completionBag, line, column, globalScope:=True)
        End Sub

        Public Shared Sub FillAllGlobalItems(
                               completionBag As CompletionBag,
                               inGlobalScope As Boolean
                         )

            If DoNotAddGlobals Then Return

            FillExpressionItems(completionBag)
            FillAllKeywords(completionBag)
            FillSubroutines(completionBag)

            If inGlobalScope Then
                FillLocals(completionBag, "")
                FillKeywords(completionBag, Token.Sub, Token.Function)
            End If
        End Sub

        Public Shared Sub FillAllKeywords(ByVal completionBag As CompletionBag)
            If DoNotAddGlobals Then Return

            FillKeywords(completionBag, Token.If, Token.For, Token.Goto, Token.While, Token.Return, Token.ExitLoop, Token.ContinueLoop)
        End Sub

        Public Shared Sub FillLogicalExpressionItems(ByVal completionBag As CompletionBag)
            If DoNotAddGlobals Then Return

            FillExpressionItems(completionBag)
            FillKeywords(completionBag, Token.And, Token.Or)
        End Sub

        Public Shared Sub FillExpressionItems(bag As CompletionBag)
            If DoNotAddGlobals Then Return

            FillTypeNames(bag)
            FillVariables(bag)
            FillBooleanLitrals(bag)
        End Sub

        Public Shared Sub FillSubroutines(bag As CompletionBag)
            If DoNotAddGlobals Then Return

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
            If DoNotAddGlobals Then Return

            For Each label In bag.SymbolTable.Labels
                bag.CompletionItems.Add(New CompletionItem() With {
                    .Name = label.Key,
                    .DisplayName = label.Value.Text,
                    .ItemType = CompletionItemType.Label
               })
            Next
        End Sub

        Public Shared Sub FillBooleanLitrals(bag As CompletionBag)
            If DoNotAddGlobals Then Return

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

        Public Shared CurrentLine As Integer
        Public Shared CurrentColumn As Integer

        Public Shared Sub FillVariables(completionBag As CompletionBag)
            If DoNotAddGlobals Then Return

            For Each variable In completionBag.SymbolTable.Variables
                Dim tokenInfo = variable.Value
                If tokenInfo.Line = CurrentLine AndAlso tokenInfo.EndColumn = CurrentColumn Then Continue For

                completionBag.CompletionItems.Add(New CompletionItem() With {
                    .Name = variable.Key,
                    .DisplayName = tokenInfo.Text,
                    .ItemType = CompletionItemType.Variable
                })
            Next
        End Sub

        Public Shared Sub FillLocals(completionBag As CompletionBag, subroutine As String)
            If DoNotAddGlobals Then Return

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
                    Dim tokenInfo = variable.Value.Identifier
                    If tokenInfo.Line = CurrentLine AndAlso tokenInfo.EndColumn = CurrentColumn Then Continue For

                    completionBag.CompletionItems.Add(New CompletionItem() With {
                        .Name = variable.Key,
                        .DisplayName = tokenInfo.Text,
                        .ItemType = CompletionItemType.Variable
                    })
                End If
            Next
        End Sub


        Public Shared Sub FillTypeNames(bag As CompletionBag)
            If DoNotAddGlobals Then Return

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

            bag.CompletionItems.Add(New CompletionItem() With {
                        .Name = "Me",
                        .DisplayName = "Me",
                        .ItemType = CompletionItemType.Keyword
            })
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
            If DoNotAddGlobals Then Return

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

        Public Sub FillDynamicMembers(bag As CompletionBag, typeName As String)
            Dim x = TrimData(typeName)

            For Each type In bag.SymbolTable.Dynamics
                Dim y = TrimData(type.Key)
                If x.Contains(y) Then
                    FillDynamicMembers(bag, type.Value)
                End If
            Next

        End Sub

        Dim numChars() As Char = {"_"c, "0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c}

        Function TrimData(name As String) As String
            If name.StartsWith("data") Then
                name = name.Substring(4)
            ElseIf name.EndsWith("data") Then
                name = name.Substring(0, name.Length - 4)
            End If
            Return name.TrimEnd(numChars).TrimStart("_"c)
        End Function

        Public Sub FillDynamicMembers(bag As CompletionBag, members As Dictionary(Of String, TokenInfo))

            For Each [property] In members.Values
                Dim Found = False
                For Each item In bag.CompletionItems
                    If item.Name = [property].Text Then
                        Found = True
                        Exit For
                    End If
                Next

                If Found Then Continue For

                bag.CompletionItems.Add(New CompletionItem() With {
                        .Name = [property].Text,
                        .DisplayName = [property].Text,
                        .ItemType = CompletionItemType.PropertyName
                 })
            Next

        End Sub
    End Class
End Namespace
