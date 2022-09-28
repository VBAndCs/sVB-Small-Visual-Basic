Imports System.IO
Imports System.Reflection
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.Completion
    Public Class CompletionHelper
        Private _compiler As Compiler
        Public Shared DoNotAddGlobals As Boolean

        Public Sub New()
            _compiler = New Compiler()
        End Sub

        Public Function GetInferedType(identifier As Token) As String
            Select Case identifier.Type
                Case TokenType.StringLiteral
                    Return NameOf(WinForms.TextEx)

                Case TokenType.DateLiteral
                    Return NameOf(WinForms.DateEx)

                Case TokenType.NumericLiteral
                    Return NameOf(WinForms.MathEx)

                Case TokenType.RightCurlyBracket
                    Return NameOf(WinForms.ArrayEx)

                Case Else
                    Dim symbolTable = _compiler.Parser.SymbolTable
                    identifier.Parent = GetStatement(_compiler, identifier.Line)
                    Dim varType = symbolTable.GetInferedType(identifier)
                    Return WinForms.PreCompiler.GetTypeName(varType)
            End Select
        End Function

        Public Function GetEmptyCompletionBag() As CompletionBag
            Return New CompletionBag(
                _compiler.TypeInfoBag,
                _compiler.Parser.SymbolTable,
                _compiler.Parser.ParseTree
            )
        End Function

        Dim completionBagCashe As New Dictionary(Of Integer, CompletionBag)

        Public Shared Function GetSubroutine(displayName As String, parseTree As List(Of Statements.Statement)) As SubroutineStatement

            If parseTree Is Nothing Then Return Nothing

            displayName = displayName.ToLower()

            For i = parseTree.Count - 1 To 0 Step -1
                Dim statement = TryCast(parseTree(i), SubroutineStatement)
                If statement IsNot Nothing AndAlso
                        statement.Name.LCaseText = displayName Then
                    Return statement
                End If
            Next

            Return Nothing
        End Function

        Public Function GetCompletionItems(
                         line As Integer,
                         column As Integer,
                         nextToEquals As Boolean,
                         nextToOperator As Boolean,
                         forHelp As Boolean
                   ) As CompletionBag

            If completionBagCashe.ContainsKey(line) Then
                Return completionBagCashe(line)
            End If

            Dim bag = GetEmptyCompletionBag()
            bag.ForHelp = forHelp
            bag.NextToEquals = nextToEquals
            bag.NextToOperator = nextToOperator

            Dim statement = GetStatement(_compiler, line)
            If statement IsNot Nothing Then
                FillCompletionItemsFromStatement(statement, bag, line, column)
                bag.SubroutineName = statement.StartToken.SubroutineName
            Else
                FillAllGlobalItems(bag, inGlobalScope:=True)
            End If

            completionBagCashe.Add(line, bag)
            Return bag
        End Function

        Public Shared Function GetStatement(
                        compiler As Compiler,
                        line As Integer
                  ) As Statement

            Return GetStatement(compiler.Parser, line)
        End Function

        Public Function GetStatement(line As Integer) As Statement
            Return GetStatement(_compiler.Parser, line)
        End Function

        Public Shared Function GetStatement(
                        parser As Parser,
                        line As Integer
                  ) As Statement

            Dim parseTree = parser.ParseTree
            If parseTree Is Nothing Then Return Nothing

            For i = parseTree.Count - 1 To 0 Step -1
                Dim statement = parseTree(i)

                If line >= statement.StartToken.Line Then
                    Return statement
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

            If inGlobalScope Then
                FillLocals(completionBag, "")
                FillKeywords(completionBag, {TokenType.Sub, TokenType.Function})
            End If

            FillExpressionItems(completionBag)
            FillAllKeywords(completionBag)
            FillSubroutines(completionBag)

        End Sub

        Private Shared allKeywords As List(Of CompletionItem)

        Public Shared Sub FillAllKeywords(completionBag As CompletionBag)
            If DoNotAddGlobals Then Return

            If AllKeywords Is Nothing Then
                AllKeywords = New List(Of CompletionItem)
                FillKeywords(
                        completionBag, {
                                TokenType.If,
                                TokenType.Goto,
                                TokenType.For,
                                TokenType.ForEach,
                                TokenType.While,
                                TokenType.Return,
                                TokenType.ExitLoop,
                                TokenType.ContinueLoop
                        }, True
                      )
            Else
                completionBag.CompletionItems.AddRange(AllKeywords)
            End If
        End Sub


        Public Shared Sub FillLogicalExpressionItems(bag As CompletionBag)
            If DoNotAddGlobals Then Return

            FillExpressionItems(bag)
        End Sub

        Public Shared Sub FillExpressionItems(bag As CompletionBag)
            If DoNotAddGlobals Then Return

            FillVariables(bag)
            FillTypeNames(bag)
            If Not bag.IsFirstToken Then FillBooleanLitrals(bag)
        End Sub

        Public Shared Sub FillSubroutines(bag As CompletionBag, Optional functionsOnly As Boolean = False)
            If DoNotAddGlobals Then Return

            For Each item In bag.SymbolTable.Subroutines
                Dim nameToken = item.Value
                If functionsOnly AndAlso nameToken.Type = TokenType.Sub Then Continue For

                Dim subName = nameToken.Text
                Dim subroutine = CompletionHelper.GetSubroutine(item.Key, bag.ParseTree)
                bag.CompletionItems.Add(New CompletionItem() With {
                    .Key = item.Key,
                    .DisplayName = subName,
                    .ReplacementText = subName & If(subroutine.Params?.Count > 0, "(", "()"),
                    .ItemType = CompletionItemType.SubroutineName,
                    .DefinitionIdintifier = nameToken
                })
            Next
        End Sub


        Public Shared Sub FillLabels(bag As CompletionBag)
            If DoNotAddGlobals Then Return

            For Each label In bag.SymbolTable.Labels
                bag.CompletionItems.Add(New CompletionItem() With {
                    .Key = label.Key,
                    .DisplayName = label.Value.Text,
                    .ItemType = CompletionItemType.Label,
                    .DefinitionIdintifier = label.Value
               })
            Next
        End Sub

        Private Shared booleanLitrals() As CompletionItem = {
                New CompletionItem() With {
                    .Key = "True",
                    .DisplayName = "True",
                    .ItemType = CompletionItemType.Keyword
                },
                New CompletionItem() With {
                    .Key = "False",
                    .DisplayName = "False",
                    .ItemType = CompletionItemType.Keyword
                },
                New CompletionItem() With {
                    .Key = "Or",
                    .DisplayName = "Or",
                    .ItemType = CompletionItemType.Keyword
                },
                New CompletionItem() With {
                    .Key = "And",
                    .DisplayName = "And",
                    .ItemType = CompletionItemType.Keyword
                }
        }

        Public Shared Sub FillBooleanLitrals(bag As CompletionBag)
            If DoNotAddGlobals Then Return

            bag.CompletionItems.Add(BooleanLitrals(0)) ' True
            bag.CompletionItems.Add(booleanLitrals(1)) ' False

            If Not bag.NextToOperator Then
                bag.CompletionItems.Add(booleanLitrals(2)) ' Or
                bag.CompletionItems.Add(booleanLitrals(3)) ' And
            End If
        End Sub


        Public Shared CurrentLine As Integer
        Public Shared CurrentColumn As Integer

        Public Shared Sub FillVariables(bag As CompletionBag)
            If DoNotAddGlobals Then Return

            For Each variable In bag.SymbolTable.GlobalVariables
                Dim varName = variable.Value

                ' Prevent shwoing new var names while typing
                If Not bag.ForHelp AndAlso varName.Line = CurrentLine AndAlso varName.EndColumn = CurrentColumn Then Continue For

                bag.CompletionItems.Add(New CompletionItem() With {
                    .Key = variable.Key,
                    .DisplayName = varName.Text,
                    .ItemType = CompletionItemType.GlobalVariable,
                    .DefinitionIdintifier = varName
                })
            Next
        End Sub

        Public Shared Sub FillLocals(bag As CompletionBag, subroutine As String)
            If DoNotAddGlobals Then Return
            Dim prfx = subroutine & "."

            For Each variable In bag.SymbolTable.LocalVariables
                If subroutine = "" Then
                    If variable.Key.Contains(".") Then Continue For
                ElseIf Not variable.Key.StartsWith(prfx) Then
                    Continue For
                End If

                Dim varName = variable.Value.Identifier
                If Not bag.ForHelp AndAlso varName.Line = CurrentLine AndAlso varName.EndColumn = CurrentColumn Then Continue For

                bag.CompletionItems.Add(New CompletionItem() With {
                        .Key = variable.Key,
                        .DisplayName = varName.Text,
                        .ItemType = CompletionItemType.LocalVariable,
                        .DefinitionIdintifier = varName
                    })
            Next
        End Sub

        Private Shared typeCompletionItems As List(Of CompletionItem)
        Private Shared membersCompletionItems As New Dictionary(Of String, List(Of CompletionItem))

        Public Shared Sub FillTypeNames(bag As CompletionBag)
            If DoNotAddGlobals Then Return

            If typeCompletionItems Is Nothing Then
                typeCompletionItems = New List(Of CompletionItem)

                For Each type In bag.TypeInfoBag.Types
                    If Not type.Value.HideFromIntellisense Then
                        typeCompletionItems.Add(New CompletionItem() With {
                            .Key = type.Key,
                            .DisplayName = type.Value.Type.Name,
                            .ItemType = CompletionItemType.TypeName,
                            .MemberInfo = type.Value.Type
                        })
                    End If
                Next

                typeCompletionItems.Add(New CompletionItem() With {
                        .ObjectName = "Form",
                        .Key = "Me",
                        .DisplayName = "Me",
                        .ItemType = CompletionItemType.Control,
                        .DefinitionIdintifier = New Token With {.Line = -1, .Type = TokenType.Identifier}
                 })
            End If

            bag.CompletionItems.AddRange(typeCompletionItems)

        End Sub

        Public Shared Sub FillMemberNames(
                        completionBag As CompletionBag,
                        typeInfo As TypeInfo,
                        objName As String
                   )

            Dim typeName = typeInfo.Type.Name.ToLower()
            If Not membersCompletionItems.ContainsKey(typeName) Then
                Dim members As New List(Of CompletionItem)

                For Each method In typeInfo.Methods
                    If Not IsHiddenFromIntellisense(method.Value) Then
                        Dim completionItem As New CompletionItem() With {
                                .Key = method.Key,
                                .DisplayName = method.Value.Name,
                                .ObjectName = objName,
                                .ItemType = CompletionItemType.MethodName,
                                .MemberInfo = method.Value,
                                .ReplacementText = method.Value.Name & "("
                         }

                        If method.Value.GetParameters().Length = 0 Then
                            completionItem.ReplacementText &= ")"
                        End If

                        members.Add(completionItem)
                    End If
                Next

                For Each [property] In typeInfo.Properties
                    If Not IsHiddenFromIntellisense([property].Value) Then
                        members.Add(New CompletionItem() With {
                                .Key = [property].Key,
                                .DisplayName = [property].Value.Name,
                                .ObjectName = objName,
                                .ItemType = CompletionItemType.PropertyName,
                                .MemberInfo = [property].Value
                         })
                    End If
                Next

                For Each [event] In typeInfo.Events
                    If Not IsHiddenFromIntellisense([event].Value) Then
                        members.Add(New CompletionItem() With {
                            .Key = [event].Key,
                            .DisplayName = [event].Value.Name,
                            .ObjectName = objName,
                            .ItemType = CompletionItemType.EventName,
                            .MemberInfo = [event].Value
                        })
                    End If
                Next

                membersCompletionItems(typeName) = members
                completionBag.CompletionItems.AddRange(members)

            Else
                completionBag.CompletionItems.AddRange(membersCompletionItems(typeName))
            End If


        End Sub

        Public Shared Sub FillWords(completionBag As CompletionBag, ParamArray words As String())
            For Each word In words
                completionBag.CompletionItems.Add(New CompletionItem() With {
                    .Key = word,
                    .DisplayName = word,
                    .ItemType = CompletionItemType.Keyword
                })
            Next
        End Sub

        Public Shared Sub FillKeywords(
                    completionBag As CompletionBag,
                   Optional keywords() As TokenType = Nothing,
                   Optional fillAllKeywords As Boolean = False
                )

            If DoNotAddGlobals Then Return

            For Each token In keywords
                Dim item = New CompletionItem() With {
                    .Key = token.ToString(),
                    .DisplayName = token.ToString(),
                    .ItemType = CompletionItemType.Keyword
                }
                completionBag.CompletionItems.Add(item)
                If fillAllKeywords Then AllKeywords.Add(item)
            Next
        End Sub

        Private Shared Function IsHiddenFromIntellisense(mi As MemberInfo) As Boolean
            Return mi.GetCustomAttributes(GetType(HideFromIntellisenseAttribute), inherit:=False).Length > 0
        End Function

        Public Sub FillDynamicMembers(bag As CompletionBag, typeName As String)
            Dim x = TrimData(typeName)

            ' Add exact obj first to get correct help info
            For Each type In bag.SymbolTable.Dynamics
                Dim y = TrimData(type.Key)
                If x = y Then
                    FillDynamicMembers(bag, typeName, type.Value)
                    Exit For
                End If
            Next

            ' Then add similar objects to add more completion properties
            For Each type In bag.SymbolTable.Dynamics
                Dim y = TrimData(type.Key)
                If x = y Then Continue For ' Add before

                If x.Contains(y) Then
                    FillDynamicMembers(bag, type.Key, type.Value)
                End If
            Next

        End Sub

        Private Shared numChars() As Char = {"_"c, "0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c}

        Public Shared Function TrimData(name As String) As String
            name = name.ToLower()
            If name.StartsWith("data") Then
                name = name.Substring(4)
            ElseIf name.EndsWith("data") Then
                name = name.Substring(0, name.Length - 4)
            End If
            Return name.TrimEnd(numChars).TrimStart("_"c)
        End Function

        Public Sub FillDynamicMembers(bag As CompletionBag, typeName As String, members As Dictionary(Of String, Token))

            For Each prop In members.Values
                ' Avoid showing new property names in completion while being typed
                If Not bag.ForHelp AndAlso prop.Line = CurrentLine AndAlso prop.EndColumn = CurrentColumn Then Continue For

                Dim Found = False
                For Each item In bag.CompletionItems
                    If item.Key = prop.Text Then
                        Found = True
                        Exit For
                    End If
                Next

                If Found Then Continue For

                bag.CompletionItems.Add(New CompletionItem() With {
                        .Key = prop.Text,
                        .DisplayName = prop.Text,
                        .ObjectName = typeName,
                        .ItemType = CompletionItemType.DynamicPropertyName,
                        .DefinitionIdintifier = prop
                 })
            Next

        End Sub

        Public Function Compile(source As TextReader, controlNames As List(Of String), moduleNames As Dictionary(Of String, String)) As Compiler
            _compiler.Parser.SymbolTable.ControlNames = controlNames
            _compiler.Parser.SymbolTable.ModuleNames = moduleNames

            _compiler.Compile(source, True)
            completionBagCashe.Clear()
            Return _compiler
        End Function
    End Class
End Namespace
