Imports System.Windows.Threading
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations
Imports Microsoft.SmallVisualBasic.Completion
Imports System.Runtime.InteropServices

Namespace Microsoft.SmallVisualBasic.LanguageService
    Public NotInheritable Class CompletionProvider
        Implements IAdornmentProvider

        Const NULL = ChrW(0)
        Private textBuffer As ITextBuffer
        Private _document As Object
        Friend textView As ITextView
        Private compHelper As CompletionHelper
        Friend adornment As CompletionAdornment
        Private dismissedSpan As ITextSpan
        Private helpUpdateTimer As DispatcherTimer
        Private editorOperations As IEditorOperations
        Private undoHistory As UndoHistory
        Public Event AdornmentsChanged As EventHandler(Of AdornmentsChangedEventArgs) Implements IAdornmentProvider.AdornmentsChanged
        Friend IsSpectialListVisible As Boolean

        Public Sub New(textView As ITextView,
                       editorOperationsProvider As IEditorOperationsProvider,
                       undoHistoryRegistry As IUndoHistoryRegistry
                   )

            Me.textView = textView
            textBuffer = textView.TextBuffer
            AddHandler textBuffer.Changed, AddressOf OnTextBufferChanged
            WordHighlightColor = textBuffer.Properties.GetProperty(Of Media.Color)("WordHighlightColor")

            editorOperations = editorOperationsProvider.GetEditorOperations(textView)
            undoHistory = undoHistoryRegistry.GetHistory(textView.TextBuffer)
            helpUpdateTimer = New DispatcherTimer(
                    TimeSpan.FromMilliseconds(500.0),
                    DispatcherPriority.ApplicationIdle,
                    AddressOf OnHelpUpdate, Application.Current.Dispatcher)

            helpUpdateTimer.Stop()
            AddHandler Me.textView.LayoutChanged, AddressOf OnLayoutChanged

            If Not textBuffer.Properties.TryGetProperty(GetType(CompletionHelper), compHelper) Then
                compHelper = New CompletionHelper()
                textBuffer.Properties.AddProperty(GetType(CompletionHelper), compHelper)
            End If
        End Sub

        Private Shared _fontNames As New List(Of CompletionItem)
        ReadOnly Property FontNames As List(Of CompletionItem)
            Get
                If _fontNames.Count = 0 Then
                    For Each family In Fonts.SystemFontFamilies
                        _fontNames.Add(New CompletionItem() With {
                            .DisplayName = family.Source,
                            .ReplacementText = $"""{family.Source}"""
                        })
                    Next
                End If
                Return _fontNames
            End Get
        End Property

        Private Shared _colorNames As New List(Of CompletionItem)
        ReadOnly Property ColorNames As List(Of CompletionItem)
            Get
                If _colorNames.Count = 0 Then
                    Dim typeInfo As TypeInfo = Nothing

                    If Compiler.TypeInfoBag.Types.TryGetValue("colors", typeInfo) Then
                        For Each item In typeInfo.Properties
                            _fontNames.Add(New CompletionItem() With {
                                .DisplayName = item.Value.Name,
                                .ReplacementText = $"""{item.Value.Name}"""
                            })
                        Next
                    End If
                End If
                Return _fontNames
            End Get
        End Property

        ReadOnly Property FormNames As List(Of CompletionItem)
            Get
                Dim formItems As New List(Of CompletionItem)
                If _document Is Nothing Then
                    _document = textBuffer.Properties.GetProperty("Document")
                End If

                Dim forms As List(Of String) = _document.GetFormNames()

                For Each form In forms
                    formItems.Add(New CompletionItem() With {
                        .DisplayName = form,
                        .ReplacementText = $"""{form}"""
                    })
                Next
                Return formItems
            End Get
        End Property

        Public Property GlobalModuleHasChanged As Boolean
            Get
                If _document Is Nothing Then
                    _document = textBuffer.Properties.GetProperty("Document")
                End If
                Return _document.GlobalModuleHasChanged
            End Get

            Set(value As Boolean)
                If _document Is Nothing Then
                    _document = textBuffer.Properties.GetProperty("Document")
                End If
                _document.GlobalModuleHasChanged = value
            End Set
        End Property

        Private ReadOnly Property GlobalParser As Parser
            Get
                If _document Is Nothing Then
                    _document = textBuffer.Properties.GetProperty("Document")
                End If

                Dim parsers As List(Of Parser) = _document.CompileGlobalModule()
                If parsers Is Nothing OrElse parsers.Count = 0 Then
                    Return Nothing
                Else
                    Return parsers(0)
                End If
            End Get
        End Property

        Public Sub CommitItem(itemWrapper As CompletionItemWrapper)
            Dim item = itemWrapper.CompletionItem
            If adornment IsNot Nothing Then
                Dim key = item.GetHistoryKey()
                If key <> "" Then
                    Dim properties = textBuffer.Properties
                    Dim controls = properties.GetProperty(Of Dictionary(Of String, String))("ControlsInfo")
                    If controls?.ContainsKey(key) Then key = controls(key)
                    CompletionHelper.History(key) = item.DisplayName
                End If

                Dim repWith = itemWrapper.ReplacementText
                Dim replaceSpan = GetReplacementSpane()

                If replaceSpan.Length = 0 And replaceSpan.Start > 0 Then
                    Select Case textView.TextSnapshot(replaceSpan.Start - 1)
                        Case "=", "<", ">", "+", "-", "*", "/"
                            repWith = " " + repWith
                    End Select
                End If

                If repWith.EndsWith("(") OrElse repWith.EndsWith(")") Then
                    Dim pos = textView.Caret.Position.TextInsertionIndex
                    Dim line = textView.TextSnapshot.GetLineFromPosition(pos)
                    Dim txt = line.GetText().Substring(pos - line.Start)
                    If LineScanner.GetFirstToken(txt, 0).Type = TokenType.LeftParens Then
                        repWith = repWith.TrimEnd("("c, ")"c)
                    End If
                End If

                If textView.TextSnapshot.GetText(replaceSpan) <> repWith Then
                    editorOperations.ReplaceText(replaceSpan, repWith, undoHistory)
                End If

                DismissAdornment(force:=True)
                    TryCast(textView, Control)?.Focus()
                    If repWith.EndsWith("(") Then ShowHelp()
                End If
        End Sub

        Friend Function GetReplacementSpane() As SnapshotSpan
            If adornment Is Nothing Then Return New SnapshotSpan(textView.TextSnapshot, textView.Caret.Position.TextInsertionIndex, 0)

            Dim replaceSpan = adornment.ReplaceSpan
            Dim snapshot = textView.TextSnapshot
            Dim span = replaceSpan.GetSpan(snapshot)
            Dim start As Integer = span.Start
            Dim pos = textView.Caret.Position.TextInsertionIndex
            If pos > span.End Then Return New SnapshotSpan(snapshot, pos, 0)

            Dim text = replaceSpan.GetText(replaceSpan.TextBuffer.CurrentSnapshot)

            Dim newStart = start
            Dim tokens = LineScanner.GetTokens(text, 0)
            Dim n = tokens.Count - 1
            Dim startIndex = 0

            For i = 0 To n
                If tokens(i).ParseType = ParseType.Operator Then
                    If tokens(i).Type = TokenType.Or OrElse tokens(i).Type = TokenType.And Then
                        Exit For
                    ElseIf i < n Then
                        newStart = start + tokens(i + 1).Column
                        startIndex = i
                    Else
                        newStart = start + tokens(i).EndColumn
                        startIndex = i
                    End If
                Else
                    Exit For
                End If
            Next

            Dim [end] = span.End
            For i = n To startIndex + 1 Step -1
                If tokens(i).ParseType = ParseType.Operator Then
                    If tokens(i).Type = TokenType.Or OrElse tokens(i).Type = TokenType.And Then
                        Exit For
                    ElseIf i > 0 Then
                        [end] = start + tokens(i - 1).EndColumn
                    Else
                        [end] = start + tokens(i).Column
                    End If
                Else
                    Exit For
                End If
            Next

            Return New SnapshotSpan(snapshot, newStart, [end] - newStart)
        End Function

        Public Sub DismissAdornment(force As Boolean)
            If adornment Is Nothing Then Return

            dismissedSpan = adornment.Span
            Dim span = dismissedSpan.GetSpan(textView.TextSnapshot)

            If span.Length = 0 AndAlso span.Start > 0 Then
                dismissedSpan = New TextSpan(textView.TextSnapshot, span.Start - 1, 1, SpanTrackingMode.EdgeInclusive)
            End If

            adornment = Nothing

            If AdornmentsChangedEvent IsNot Nothing Then
                Application.Current.Dispatcher.Invoke(
                    DispatcherPriority.Normal,
                    CType(Function()
                              RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(dismissedSpan))
                              Return Nothing
                          End Function, DispatcherOperationCallback),
                    Nothing)
            End If

            If Not force Then dismissedSpan = Nothing
        End Sub

        Public Function GetAdornments(span As SnapshotSpan) As IList(Of IAdornment) Implements IAdornmentProvider.GetAdornments
            Dim list As New List(Of IAdornment)()

            If adornment IsNot Nothing AndAlso adornment.Span.GetSpan(span.Snapshot).OverlapsWith(span) Then
                list.Add(adornment)
            End If

            Return list
        End Function

        Private Sub OnLayoutChanged(sender As Object, e As TextViewLayoutChangedEventArgs)
            AddHandler textView.Caret.PositionChanged, AddressOf OnCaretPositionChanged
            RemoveHandler textView.LayoutChanged, AddressOf OnLayoutChanged
        End Sub

        Public Shared popHelp As Primitives.Popup

        Private Sub OnCaretPositionChanged(sender As Object, e As CaretPositionChangedEventArgs)
            If adornment IsNot Nothing Then
                Dim span = adornment.Span.GetSpan(textView.TextSnapshot)
                Dim textInsertionIndex = e.NewPosition.TextInsertionIndex

                If textInsertionIndex < span.Start OrElse textInsertionIndex > span.End Then
                    DismissAdornment(force:=False)
                End If
            Else
                helpUpdateTimer.Stop()
                popHelp.IsOpen = False
                helpUpdateTimer.Start()
            End If
        End Sub

        Dim lastSpan As Span
        Dim helpCashe As New Dictionary(Of Span, CompletionItemWrapper)
        Dim highlightCashe As New Dictionary(Of Span, Span())
        Dim callerCashe As New Dictionary(Of Span, CallerInfo)

        Private Sub OnHelpUpdate(sender As Object, e As EventArgs)
            helpUpdateTimer.Stop()
            IsSpectialListVisible = False

            Dim snapshot = textBuffer.CurrentSnapshot
            If snapshot.Length = 0 Then Return

            Dim pos = textView.Caret.Position.TextInsertionIndex
            Dim line = snapshot.GetLineFromPosition(pos)
            Dim column = pos - line.Start
            Dim paramIndex As Integer = -1

            Dim currentLine = line.LineNumber
            Dim startLine = If(currentLine = 0, 0, currentLine - 1)
            Dim endLine = If(currentLine < snapshot.LineCount - 1, currentLine + 1, currentLine)

            Dim tokens = GetTokens(snapshot, startLine, currentLine, endLine)
            If tokens.Count = 0 Then Return

            endLine = startLine + tokens.Last.Line

            Dim force = False
            Dim symbol = ""
            Dim checkEspecialArgs = False

            If TypeOf sender Is Boolean Then
                force = CType(sender, Boolean)
            ElseIf TypeOf sender Is String OrElse TypeOf sender Is Char Then
                symbol = CType(sender, String)
            ElseIf TypeOf sender Is (String, Boolean) Then
                Dim info = CType(sender, (Symbol As String, Force As Boolean))
                force = info.Force
                symbol = info.Symbol
                If symbol = ", " Then
                    symbol = ","
                    column -= 2
                ElseIf symbol = "*" Then
                    symbol = ""
                    checkEspecialArgs = True
                End If
            End If

            Dim span As New Span(pos, 0)

            If Not ParseSepTokens(
                    tokens, startLine, currentLine, endLine,
                    symbol, force, line, column, span,
                    paramIndex) Then Return

            If sourceCodeChanged OrElse GlobalModuleHasChanged Then
                helpCashe.Clear()
                highlightCashe.Clear()

            ElseIf helpCashe.ContainsKey(span) Then
                Dim editor = CType(textView, AvalonTextView).Editor
                If Not editor.ContainsWordHighlights Then
                    editor.HighlightWords(WordHighlightColor, highlightCashe(span))
                End If

                Dim wrapper = helpCashe(span)
                If wrapper IsNot Nothing Then
                    wrapper.CompletionItem.ParamIndex = paramIndex
                    ShowPopupHelp(wrapper)
                End If
                lastSpan = span
                If Not checkEspecialArgs Then Return
            End If

            If symbol <> "" AndAlso symbol <> "(" AndAlso line.Start + column = pos Then Return
            ShowHelpInfo(
                line,
                column,
                paramIndex,
                span,
                If(checkEspecialArgs OrElse force, ",", symbol)
            )

        End Sub

        Private Function ParseSepTokens(
                    tokens As List(Of Token),
                    startLine As Integer,
                    currentLine As Integer,
                    endLine As Integer,
                    symbol As String,
                    force As Boolean,
                    ByRef line As ITextSnapshotLine,
                    ByRef column As Integer,
                    <Out> ByRef span As Span,
                    ByRef paramIndex As Integer
                ) As Boolean

            Dim tempLine = line
            Dim tempColumn = column
            Dim snapshot = textView.TextSnapshot
            Dim prevChar = NULL
            Dim currentChar = NULL
            Dim prevIsSep = False
            Dim currentToken = Token.Illegal
            Dim exactToken = Token.Illegal

            If symbol = "" Then
                Dim prevToken As Token
                Dim nextToken As Token
                Dim n = tokens.Count - 1
                Dim notFound = True

                For i = 0 To n
                    Dim token = tokens(i)
                    If token.Contains(currentLine, column, Not textView.Selection.IsEmpty) Then
                        prevToken = GetNonCommentToken(tokens, i - 1, True)
                        prevIsSep = IsPrevSeparator(currentLine, symbol, prevToken, prevChar)
                        exactToken = token

                        If column > token.Column Then
                            currentToken = If(
                                  prevIsSep AndAlso token.ParseType = ParseType.Comment,
                                  Token.Illegal,
                                  token
                            )
                            nextToken = GetNonCommentToken(tokens, i + 1, False)

                        ElseIf Not (prevIsSep AndAlso textView.Selection.IsEmpty) Then
                            currentToken = token
                            nextToken = GetNonCommentToken(tokens, i + 1, False)
                        End If

                        notFound = False
                        Exit For

                    ElseIf token.Line > currentLine Then
                        prevToken = GetNonCommentToken(tokens, i - 1, True)
                        prevIsSep = IsPrevSeparator(currentLine, symbol, prevToken, prevChar)
                        nextToken = GetNonCommentToken(tokens, i, False)
                        notFound = False
                        Exit For

                    ElseIf token.Line = currentLine Then
                        If token.EndColumn = column Then ' pos is right after the end of the current token
                            exactToken = token
                            prevToken = GetNonCommentToken(tokens, i, True)
                            prevIsSep = IsPrevSeparator(currentLine, symbol, prevToken, prevChar)
                            nextToken = GetNonCommentToken(tokens, i + 1, True)
                            notFound = False
                            Exit For

                        ElseIf token.Column > column Then
                            prevToken = GetNonCommentToken(tokens, i - 1, True)
                            prevIsSep = IsPrevSeparator(currentLine, symbol, prevToken, prevChar)
                            notFound = False
                            Exit For

                        ElseIf i = n Then
                            prevToken = GetNonCommentToken(tokens, i, True)
                            prevIsSep = IsPrevSeparator(currentLine, symbol, prevToken, prevChar)
                            notFound = False
                            Exit For
                        End If
                    End If
                Next

                If notFound Then
                    prevToken = GetNonCommentToken(tokens, n, True)
                    prevIsSep = IsPrevSeparator(currentLine, symbol, prevToken, prevChar)
                End If

                span = If(currentToken.IsIllegal,
                     GetSpan(prevToken, nextToken, snapshot, startLine, endLine),
                     GetSpan(currentToken, snapshot, startLine)
                 )

                If Not sourceCodeChanged AndAlso span = lastSpan AndAlso Not force Then
                    Dim editor = CType(textView, AvalonTextView).Editor
                    If Not editor.ContainsWordHighlights Then
                        editor.HighlightWords(WordHighlightColor, highlightCashe(lastSpan))
                    End If
                    Return False
                End If

                If prevIsSep Then
                    If (currentToken.IsIllegal OrElse currentToken.Type = TokenType.NumericLiteral OrElse currentToken.Type = TokenType.StringLiteral) AndAlso symbol = "" Then
                        line = snapshot.GetLineFromLineNumber(prevToken.Line + startLine)
                        column = prevToken.EndColumn
                    End If

                ElseIf symbol = "" Then
                    Dim validToken = If(currentToken.IsIllegal, nextToken, currentToken)
                    Select Case validToken.Type
                        Case TokenType.Comma, TokenType.RightParens
                            currentChar = validToken.Text(0)
                            If currentToken.IsIllegal AndAlso symbol = "" Then
                                line = snapshot.GetLineFromLineNumber(validToken.Line + startLine)
                                column = validToken.Column
                            End If
                        Case Else
                            currentChar = NULL
                    End Select

                Else
                    span = New Span(textView.Caret.Position.TextInsertionIndex, 1)
                End If
            End If

            Dim isClosing = currentChar = ")" OrElse symbol = ")"
            Dim isComma = currentChar = "," OrElse symbol = ","
            Dim commaOrParans = prevIsSep OrElse isClosing OrElse isComma

            If commaOrParans OrElse symbol = "(" Then
                Dim editor = CType(textView, AvalonTextView).Editor
                Dim caller As CallerInfo

                If sourceCodeChanged Then
                    callerCashe.Clear()
                    caller = Parser.GetCommaCallerInfo(
                            editor.Text,
                            line.LineNumber,
                            column - If(prevIsSep, 1, 0)
                    )
                    callerCashe(span) = caller

                ElseIf callerCashe.ContainsKey(span) Then
                    caller = callerCashe(span)

                Else
                    caller = Parser.GetCommaCallerInfo(editor.Text, line.LineNumber, column - If(prevIsSep, 1, 0))
                    callerCashe(span) = caller
                End If

                If caller Is Nothing Then
                    line = tempLine
                    column = tempColumn
                    If symbol = "" AndAlso exactToken.IsIllegal Then
                        Return False
                    End If

                Else
                    line = textView.TextSnapshot.GetLineFromLineNumber(caller.Line)
                    column = caller.EndColumn
                    paramIndex = caller.ParamIndex -
                         If(currentChar = ",", 1, 0) +
                         If(prevChar = ")" OrElse symbol = ")", 1, 0)
                End If

            ElseIf prevChar = "("c Then
                paramIndex = 0
                column -= 1
            End If

            Return True
        End Function

        Private Sub ShowHelpInfo(
                       line As ITextSnapshotLine,
                       column As Integer,
                       paramIndex As Integer,
                       span As Span,
                       symbol As String
             )

            Dim currentToken As Token
            Dim bag = GetCompletionBag(line, column, currentToken, True)

            sourceCodeChanged = False
            Dim editor = CType(textView, AvalonTextView).Editor

            If bag Is Nothing OrElse currentToken.IsIllegal OrElse
                        bag.CompletionItems.Count <= 0 Then
                highlightCashe(span) = Nothing
                helpCashe(span) = Nothing
                lastSpan = span
                Return
            End If

            lastSpan = span

            Dim tokenText = currentToken.LCaseText
            For Each item In bag.CompletionItems
                If item.Key = "##" Then
                    Dim wrapper = New CompletionItemWrapper(item, bag)
                    ShowPopupHelp(wrapper)
                    helpCashe(span) = wrapper
                    highlightCashe(span) = Nothing
                    Return

                ElseIf item.DisplayName?.ToLower() = tokenText Then
                    If bag.IsMethod Then
                        If item.ItemType <> CompletionItemType.MethodName Then Continue For
                    End If

                    Dim spans = HeighLightIdentifiers(
                                currentToken,
                                bag.SubroutineName,
                                item.ItemType,
                                bag.SymbolTable,
                                textView.TextSnapshot
                    )?.ToArray()

                    highlightCashe(span) = spans
                    If Not editor.ContainsWordHighlights Then
                        editor.HighlightWords(WordHighlightColor, spans)
                    End If

                    item.ParamIndex = paramIndex

                    If item.ItemType = CompletionItemType.Control AndAlso item.DisplayName = "Me" Then
                        Dim formName As String
                        If textBuffer.Properties.TryGetProperty("FormName", formName) Then
                            item.Key = formName
                        End If
                    End If

                    If byConventionName <> "" Then
                        item.ObjectName = byConventionName
                    End If

                    Dim wrapper = New CompletionItemWrapper(item, bag)
                    If item.ParamIndex > -1 AndAlso
                                (symbol = "(" OrElse symbol = ",") AndAlso
                                wrapper.Documentation IsNot Nothing Then
                        Dim params = wrapper.Documentation.ParamsDoc.Keys

                        If item.ParamIndex < params.Count Then
                            Dim name = params(item.ParamIndex).ToLower().Trim("_")
                            Dim especialItem = GetEspecialItem(name)

                            If especialItem = "" AndAlso name.Length > 2 AndAlso
                                       name.StartsWith("is") AndAlso
                                       Char.IsUpper(params(item.ParamIndex)(2)
                                    ) Then
                                especialItem = "Boolean"
                            End If

                            If especialItem <> "" Then
                                Dim pos = textView.Caret.Position.TextInsertionIndex
                                line = textView.TextSnapshot.GetLineFromPosition(pos)
                                Dim txt = line.GetText().Substring(pos - line.Start)
                                Dim token = LineScanner.GetFirstToken(txt, 0)
                                IsSpectialListVisible = True

                                If especialItem.ToLower <> token.LCaseText AndAlso
                                        Not (especialItem.EndsWith("Name") AndAlso token.Type = TokenType.StringLiteral) AndAlso Not (
                                                especialItem = "Boolean" AndAlso (token.Type = TokenType.False OrElse token.Type = TokenType.True)
                                        ) Then
                                    bag = GetCompletionBag(especialItem)
                                    token.Column = pos - line.Start
                                    token.Text = ""
                                    canGetOriginalBag = True
                                    ShowCompletionAdornment(
                                        textView.TextSnapshot,
                                        bag,
                                        token,
                                        line
                                    )
                                End If
                            End If
                        End If
                    End If

                    ShowPopupHelp(wrapper)
                    helpCashe(span) = wrapper
                    Return
                End If
            Next

            highlightCashe(span) = Nothing
            helpCashe(span) = Nothing
        End Sub

        Private Function HeighLightIdentifiers(
                            currentToken As Token,
                            subroutineName As String,
                            itemType As CompletionItemType,
                            symbolTable As SymbolTable,
                            snapshot As ITextSnapshot
                    ) As List(Of Span)

            If Not currentToken.Type = TokenType.Identifier Then Return Nothing

            Dim idName = currentToken.LCaseText
            Dim type As CompletionItemType
            Dim subName = ""

            Select Case itemType
                Case CompletionItemType.EventName
                    type = CompletionItemType.PropertyName
                Case CompletionItemType.Control, CompletionItemType.TypeName
                    type = CompletionItemType.GlobalVariable
                Case CompletionItemType.LocalVariable, CompletionItemType.Label
                    type = itemType
                    subName = subroutineName.ToLower()
                Case Else
                    type = itemType
            End Select

            Dim spans = (
                    From id In symbolTable.AllIdentifiers
                    Where id.SymbolType = type AndAlso
                                id.LCaseText = idName AndAlso
                                (subName = "" OrElse id.SubroutineName.ToLower() = subName)
                    Let start = snapshot.GetLineFromLineNumber(id.Line).Start
                    Select New Span(start + id.Column, id.EndColumn - id.Column)
             ).ToList()

            Return spans

        End Function


        Private Function GetTokens(
                     snapshot As ITextSnapshot,
                     ByRef startLine As Integer,
                     ByRef currentLine As Integer,
                     endLine As Integer) As List(Of Token)

            Dim tokens As New List(Of Token)
            Dim lineTokens As List(Of Token)
            Dim nextLineText = snapshot.GetLineFromLineNumber(startLine).GetText()
            Dim addNextLine = True

            For i = startLine To endLine - 1
                Dim curLineText = nextLineText
                nextLineText = snapshot.GetLineFromLineNumber(i + 1).GetText()
                lineTokens = LineScanner.GetTokens(curLineText, i - startLine)
                addNextLine = LineScanner.IsLineContinuity(lineTokens, nextLineText)

                If i = currentLine OrElse addNextLine Then
                    tokens.AddRange(lineTokens)
                Else
                    startLine = currentLine 'ignore prev line
                    If startLine = endLine Then
                        addNextLine = True
                        Exit For
                    End If
                End If
            Next

            If addNextLine Then
                lineTokens = LineScanner.GetTokens(nextLineText, endLine - startLine)
                tokens.AddRange(lineTokens)
            End If

            currentLine -= startLine  '  line pos abs for scaned tokens
            Return tokens
        End Function

        Private Shared Function IsPrevSeparator(
                    currentLine As Integer,
                    symbole As String,
                    prevToken As Token,
                    ByRef prevChar As Char) As Boolean

            Select Case prevToken.Type
                Case TokenType.Comma, TokenType.LeftParens
                    prevChar = prevToken.Text(0)
                    Return True

                Case TokenType.RightParens
                    If prevToken.Line = currentLine AndAlso symbole = "" Then
                        prevChar = prevToken.Text
                        Return True
                    End If
            End Select

            Return False
        End Function

        Private Function GetSpan(
                                currentToken As Token,
                                snapshot As ITextSnapshot,
                                startLine As Integer) As Span

            Dim linePos = snapshot.GetLineFromLineNumber(startLine + currentToken.Line).Start
            Return New Span(linePos + currentToken.Column, currentToken.EndColumn - currentToken.Column)
        End Function

        Private Function GetSpan(
                                startToken As Token,
                                endToken As Token,
                                snapshot As ITextSnapshot,
                                startLine As Integer,
                                endLine As Integer
                     ) As Span

            Dim start As Integer
            If startToken.IsIllegal Then
                start = snapshot.GetLineFromLineNumber(startLine).Start
            Else
                Dim line1Pos = snapshot.GetLineFromLineNumber(startLine + startToken.Line).Start
                start = line1Pos + startToken.EndColumn
            End If

            Dim [end] As Integer
            If endToken.IsIllegal Then
                [end] = snapshot.GetLineFromLineNumber(endLine).End
            ElseIf endToken = startToken Then
                [end] = start + 1
            Else
                Dim line2Pos = snapshot.GetLineFromLineNumber(startLine + endToken.Line).Start
                [end] = line2Pos + endToken.Column
            End If

            Return New Span(start, [end] - start)
        End Function

        Private Function GetNonCommentToken(
                           tokens As List(Of Token),
                           index As Integer,
                           moveBack As Boolean
                   ) As Token

            If index < 0 OrElse index >= tokens.Count Then
                Return Token.Illegal
            End If

            If tokens(index).Type <> TokenType.Comment Then
                Return tokens(index)
            End If


            Return GetNonCommentToken(
                    tokens,
                    index + If(moveBack, -1, 1),
                    moveBack
           )
        End Function

        Dim sourceCodeChanged As Boolean = True
        Dim needsToReCompile As Boolean = True

        Private Sub OnTextBufferChanged(sender As Object, e As Nautilus.Text.TextChangedEventArgs)
            sourceCodeChanged = True
            needsToReCompile = True
            popHelp.IsOpen = False
            helpUpdateTimer.Stop()

            Dim textChange = e.Changes(0)
            Dim newText = textChange.NewText.Trim(" "c, vbTab)

            Dim pos = textView.Caret.Position.TextInsertionIndex
            If Math.Abs(pos - textChange.Position) > 1 Then
                popHelp.IsOpen = False
                DismissAdornment(force:=False)
                Return
            End If

            Dim newEnd = textChange.NewEnd

            If adornment IsNot Nothing Then
                Dim span = adornment.Span.GetSpan(e.After)

                If span.IsEmpty OrElse newEnd < span.Start OrElse
                            textChange.NewText.IndexOfAny({vbCr, vbLf}) > -1 Then
                    DismissAdornment(force:=False)
                Else
                    Dim snapshot = e.After
                    Dim line = snapshot.GetLineFromPosition(newEnd)
                    Dim tokens = LineScanner.GetTokens(line.GetText(), line.LineNumber)
                    Dim curToken As Token
                    Dim index = ParseTokens(tokens, newEnd - line.Start, curToken, Nothing, Nothing)
                    Dim adornmentSpan = GetTextSpanFromToken(line, curToken)
                    Dim replaceSpan = adornmentSpan

                    If replaceSpan.GetSpan(snapshot).IsEmpty AndAlso line.TextSnapshot.Length > 0 Then
                        If curToken.Column = 0 Then
                            adornmentSpan = New TextSpan(
                                    snapshot,
                                    line.Start,
                                    System.Math.Min(curToken.EndColumn + 1, snapshot.Length),
                                    SpanTrackingMode.EdgeInclusive
                             )
                        Else
                            adornmentSpan = New TextSpan(
                                snapshot,
                                line.Start + curToken.Column - 1,
                                System.Math.Min(curToken.EndColumn - curToken.Column + 1, snapshot.Length),
                                SpanTrackingMode.EdgeInclusive
                            )
                        End If
                    End If

                    adornment.ModifySpans(adornmentSpan, replaceSpan)
                End If

            ElseIf newText <> "" Then
                Dim c = newText.Last
                If Char.IsLetter(c) OrElse c = "_" Then
                    ShowCompletionAdornment(e.After, newEnd)
                Else Select Case c
                        Case "."c, "!"c
                            popHelp.IsOpen = False
                            ShowCompletionAdornment(e.After, newEnd)

                        Case "="c
                            ShowCompletionAdornment(e.After, newEnd, True)
                            Dim line = e.After.GetLineFromPosition(newEnd)
                            Dim compiler = compHelper.Compile(New IO.StringReader(line.GetText()), Nothing, Nothing, Nothing)
                            Dim vars = compiler.Parser.SymbolTable.GlobalVariables

                            If vars.Count > 0 Then
                                Dim varName = vars.Values(0).LCaseText
                                CompletionHelper.History("_" & varName(0)) = varName
                            End If

                        Case ">"c, "<"c
                            ShowCompletionAdornment(e.After, newEnd, True)

                        Case "("c, ")"c, ","c
                            OnHelpUpdate(c, Nothing)
                    End Select
                End If
            End If
        End Sub

        Dim byConventionName As String

        Public Function GetCompletionBag(
                          line As ITextSnapshotLine,
                          column As Integer,
                          <Out> ByRef currentToken As Token,
                          Optional forHelp As Boolean = False
                    ) As CompletionBag

            CompletionHelper.CurrentLine = line.LineNumber
            CompletionHelper.CurrentColumn = column
            byConventionName = ""

            Dim properties = textBuffer.Properties
            Dim controlsInfo = properties.GetProperty(Of Dictionary(Of String, String))("ControlsInfo")
            Dim controlNames = properties.GetProperty(Of List(Of String))("ControlNames")
            Dim tokens = LineScanner.GetTokens(line.GetText(), line.LineNumber)
            Dim lastIndex = tokens.Count - 1
            Dim prevToken As Token = Nothing
            Dim b4PrevToken As Token = Nothing
            Dim index = ParseTokens(tokens, column, currentToken, prevToken, b4PrevToken)

            Dim isFirstToken = (index < 1)
            Dim identifierToken = Token.Illegal

            Dim isLookup = (prevToken.Type = TokenType.Lookup)

            If prevToken.Type = TokenType.Dot OrElse isLookup Then
                identifierToken = b4PrevToken
                CompletionHelper.DoNotAddGlobals = True

            Else
                isLookup = (currentToken.Type = TokenType.Lookup)
                If currentToken.Type = TokenType.Dot OrElse isLookup Then
                    CompletionHelper.DoNotAddGlobals = True
                    identifierToken = prevToken
                    Dim endColumn = currentToken.EndColumn
                    currentToken = Token.Illegal
                    currentToken.Column = endColumn
                    currentToken.Text = ""
                End If
            End If

            Dim source = New TextBufferReader(line.TextSnapshot)
            If GlobalModuleHasChanged Then
                needsToReCompile = True
                GlobalModuleHasChanged = False
            End If

            Dim gp = GlobalParser
            If needsToReCompile Then
                compHelper.Compile(
                    source, controlNames,
                    controlsInfo,
                    gp
                )
                needsToReCompile = False
            End If

            Dim newBag = compHelper.GetEmptyCompletionBag(gp)
            newBag.ForHelp = forHelp
            Dim addGlobals = False

            If identifierToken.IsIllegal Then
                addGlobals = True
                If currentToken.Type = TokenType.DateLiteral Then
                    If forHelp Then
                        newBag.CompletionItems.Add(New CompletionItem() With {
                            .DisplayName = currentToken.Text,
                            .Key = "##",
                            .ItemType = CompletionItemType.Lireral
                        })
                    End If
                    Return newBag

                ElseIf currentToken.Type = TokenType.NumericLiteral Then
                    Return newBag

                ElseIf currentToken.Type = TokenType.StringLiteral Then
                    If Not forHelp AndAlso (
                                prevToken.Type = TokenType.Equals OrElse
                                prevToken.Type = TokenType.LeftParens
                            ) Then
                        Dim method = b4PrevToken.LCaseText
                        If method.Contains("fontname") Then
                            newBag.CompletionItems.AddRange(FontNames)
                        ElseIf method.Contains("showform") OrElse
                                   method.Contains("showdialog") OrElse
                                   method.Contains("showchildform") OrElse
                                   method.Contains("runformtests") Then
                            newBag.CompletionItems.AddRange(FormNames)
                        ElseIf method.Contains("color") Then
                            newBag.CompletionItems.AddRange(ColorNames)
                        End If
                    End If
                    Return newBag

                ElseIf controlsInfo IsNot Nothing Then
                    Dim txt = currentToken.LCaseText
                    If txt <> "" AndAlso controlNames IsNot Nothing Then
                        Dim x = currentToken.Text
                        Dim txt2 = UCase(x(0)) & If(x.Length > 1, x.Substring(0), "")
                        Dim controls = From name In controlNames
                                       Where name(0) <> "("c AndAlso (
                                           (forHelp AndAlso name.ToLower() = txt) OrElse
                                           (Not forHelp AndAlso (name.ToLower().StartsWith(txt) OrElse name.Contains(txt2)))
                                       )

                        If controls.Count = 0 Then
                            Dim moduleName = WinForms.PreCompiler.GetModuleFromVarName(txt)
                            If moduleName <> "" Then
                                FillMemberNames(newBag, moduleName, currentToken.Text)
                                If forHelp Then byConventionName = moduleName
                            End If

                        Else
                            For Each name In controls
                                newBag.CompletionItems.Add(
                                    New CompletionItem With {
                                        .ObjectName = controlsInfo(name.ToLower()),
                                        .DisplayName = name,
                                        .ItemType = CompletionItemType.Control,
                                        .ReplacementText = name,
                                        .DefinitionIdintifier = New Token() With {.Line = -1, .Type = TokenType.Identifier}
                                    }
                                )
                            Next
                        End If
                    End If
                End If

            Else
                Dim typeInfo As TypeInfo = Nothing
                Dim name = identifierToken.LCaseText

                If newBag.TypeInfoBag.Types.TryGetValue(name, typeInfo) Then
                    CompletionHelper.FillMemberNames(newBag, typeInfo, identifierToken.Text)

                ElseIf isLookup OrElse name.StartsWith("data") OrElse name.EndsWith("data") Then
                    compHelper.FillDynamicMembers(newBag, identifierToken.Text)

                ElseIf controlsInfo?.ContainsKey(name) Then
                    FillMemberNames(newBag, controlsInfo(name), identifierToken.Text)

                Else
                    Dim moduleName = compHelper.GetInferedType(identifierToken)
                    If moduleName = "" Then
                        moduleName = WinForms.PreCompiler.GetModuleFromVarName(name)
                    End If

                    If moduleName <> "" Then
                        FillMemberNames(newBag, moduleName, identifierToken.Text)
                        If forHelp Then
                            byConventionName = moduleName
                            newBag.IsMethod = index < lastIndex AndAlso tokens(index + 1).Type = TokenType.LeftParens
                        End If
                    End If
                End If
            End If

            If addGlobals OrElse newBag Is Nothing OrElse newBag.CompletionItems.Count = 0 Then
                If Not (currentToken.Type = TokenType.StringLiteral OrElse currentToken.Type = TokenType.Comment) Then
                    ' Fix prevToken if current token not found
                    If currentToken = Token.Illegal AndAlso index > 0 Then
                        prevToken = tokens(index - 1)
                    End If

                    Dim bag = compHelper.GetCompletionItems(
                        line.LineNumber, column,
                        prevToken.Type = TokenType.Equals OrElse currentToken.Type = TokenType.Equals,
                        IsCompletionOperator(prevToken) OrElse IsCompletionOperator(currentToken),
                        forHelp,
                        gp
                    )

                    If bag.ShowCompletion Then
                        bag.CompletionItems.AddRange(newBag.CompletionItems)
                    End If
                    newBag = bag
                End If
            End If

            If forHelp Then
                If currentToken.IsIllegal AndAlso index < lastIndex Then
                    currentToken = tokens(index + 1)
                End If
            Else
                newBag.CompletionItems.Sort(
                   Function(ci1, ci2) ci1.DisplayName.CompareTo(ci2.DisplayName))
            End If

            CompletionHelper.DoNotAddGlobals = False
            newBag.IsFirstToken = isFirstToken
            Return newBag
        End Function

        Private Shared Function ParseTokens(
                            tokens As List(Of Token),
                            column As Integer,
                            ByRef currentToken As Token,
                            ByRef prevToken As Token,
                            ByRef b4PrevToken As Token
                     ) As Integer

            Dim n = tokens.Count - 1
            currentToken = Token.Illegal
            prevToken = Token.Illegal
            b4PrevToken = Token.Illegal
            Dim index = -1

            For i = 0 To n
                Dim token = tokens(i)
                If token.Column > column Then Exit For
                If column >= token.Column AndAlso column <= token.EndColumn Then
                    If token.ParseType = ParseType.Operator Then
                        If column = token.Column AndAlso (token.Type = TokenType.RightBracket OrElse token.Type = TokenType.RightParens OrElse token.Type = TokenType.RightCurlyBracket) Then
                            Exit For
                        ElseIf i < n AndAlso tokens(i + 1).Column = token.EndColumn AndAlso tokens(i + 1).ParseType <> ParseType.Operator Then
                            b4PrevToken = currentToken
                            prevToken = token
                            currentToken = tokens(i + 1)
                            index = i + 1
                        Else
                            b4PrevToken = prevToken
                            prevToken = currentToken
                            currentToken = token
                            index = i
                        End If
                    Else
                        b4PrevToken = prevToken
                        prevToken = currentToken
                        currentToken = token
                        index = i
                    End If
                    Exit For
                End If

                If i = n Then
                    If column <= token.EndColumn Then
                        If Not (token.ParseType = ParseType.Operator AndAlso column = currentToken.EndColumn) Then
                            b4PrevToken = prevToken
                            prevToken = currentToken
                            currentToken = token
                            index = i
                        End If
                    Else
                        b4PrevToken = prevToken
                        prevToken = currentToken
                        currentToken = Token.Illegal
                        index = i + 1
                    End If

                Else
                    b4PrevToken = prevToken
                    prevToken = currentToken
                    currentToken = token
                    index = i
                End If

            Next
            Return index
        End Function

        Private Function IsCompletionOperator(token As Token) As Boolean
            ' `Return False` means we can ahow `And` and `Or`
            If token.Type = TokenType.RightBracket Then Return False
            If token.Type = TokenType.RightCurlyBracket Then Return False
            If token.Type = TokenType.RightParens Then Return False

            If token.Type = TokenType.True Then Return False
            If token.Type = TokenType.False Then Return False

            If token.ParseType = ParseType.Keyword Then Return True
            Return token.ParseType = ParseType.Operator
        End Function

        Dim ShowCompletion As New RunAction()

        Public Sub ShowCompletionAdornment(
                           snapshot As ITextSnapshot,
                           caretPosition As Integer,
                           Optional checkEspecialItem As Boolean = False,
                           Optional ctrlSpace As Boolean = False
                    )

            ' Wait until code editor respond to changes, to avoid any conflicts
            ShowCompletion.After(
                  20,
                  Sub() DoShowCompletion(snapshot, caretPosition, checkEspecialItem, ctrlSpace)
            )
        End Sub

        Private Sub DoShowCompletion(
                           snapshot As ITextSnapshot,
                           caretPosition As Integer,
                           checkEspecialItem As Boolean,
                           ctrlSpace As Boolean
                    )

            Dim bag As CompletionBag
            Dim especialItem = ""
            Dim curToken As Token

            Dim line = snapshot.GetLineFromPosition(caretPosition)
            If checkEspecialItem Then
                Dim tokens = LineScanner.GetTokens(line.GetText(), line.LineNumber)
                If tokens.Count < 2 Then
                    If ctrlSpace Then GoTo LineShow
                    Return
                End If

                Dim n = GetLastTokenIndex(line, caretPosition, tokens)
                If n = 0 Then
                    If ctrlSpace Then GoTo LineShow
                    Return
                End If

                Select Case tokens(n).Type
                    Case TokenType.Equals, TokenType.NotEqualTo,
                             TokenType.GreaterThan, TokenType.GreaterThanEqualTo,
                             TokenType.LessThan, TokenType.LessThanEqualTo

                        Dim index = n - 1
                        curToken = tokens(n)
                        Dim token = GetIdentifier(tokens, index, line)
                        Dim name = tokens(index).LCaseText.Trim("_")
                        especialItem = GetEspecialItem(name)

                        If especialItem = "" Then
                            especialItem = InferEspType(token, line)
                            If especialItem = "" AndAlso Not ctrlSpace Then
                                Return
                            End If
                        End If
                End Select
            End If

LineShow:
                If especialItem = "" Then
                bag = GetCompletionBag(line, caretPosition - line.Start, curToken)
                canGetOriginalBag = False
            Else
                bag = GetCompletionBag(especialItem)
                canGetOriginalBag = True
            End If

            If bag IsNot Nothing AndAlso bag.CompletionItems.Count > 0 Then
                bag.CtrlSpace = ctrlSpace
                ShowCompletionAdornment(snapshot, bag, curToken, line)
            End If

        End Sub

        Private Function InferEspType(token As Token, line As ITextSnapshotLine) As String
            Dim especialItem = ""
            Dim properties = textBuffer.Properties
            Dim controlsInfo = properties.GetProperty(Of Dictionary(Of String, String))("ControlsInfo")
            Dim controlNames = properties.GetProperty(Of List(Of String))("ControlNames")
            Dim source = New TextBufferReader(line.TextSnapshot)
            Dim gp = GlobalParser
            Dim symbolTable = compHelper.Compile(
                                source, controlNames, controlsInfo, gp
                            ).Parser.SymbolTable

            needsToReCompile = False
            token.Parent = compHelper.GetRootStatement(line.LineNumber)
            Dim varType = symbolTable.GetInferedType(token)

            If varType = VariableType.Any Then
                Dim st = TryCast(token.Parent.GetStatementAt(line.LineNumber), Statements.AssignmentStatement)
                If st IsNot Nothing Then
                    Dim prop = TryCast(st.LeftValue, Expressions.PropertyExpression)
                    If prop Is Nothing OrElse Not prop.IsEvent Then
                        varType = If(st.LeftValue Is Nothing, VariableType.Any, st.LeftValue.InferType(symbolTable))
                    Else
                        especialItem = "Sub"
                    End If
                End If
            End If

            Select Case varType
                Case VariableType.Color
                    especialItem = "Colors"
                Case VariableType.Key
                    especialItem = "Keys"
                Case VariableType.DialogResult
                    especialItem = "DialogResults"
                Case VariableType.ControlType
                    especialItem = "ControlTypes"
                Case VariableType.Boolean
                    especialItem = "Boolean"
            End Select

            Return especialItem
        End Function

        Private Function GetEspecialItem(name As String) As String
            If name.StartsWith("color") OrElse name.EndsWith("color") OrElse name.EndsWith("colors") Then
                Return "Colors"
            ElseIf name.StartsWith("key") OrElse name.EndsWith("key") OrElse name.EndsWith("keys") Then
                Return "Keys"
            ElseIf name = "showdialog" OrElse name.Contains("dialogresult") Then
                Return "DialogResults"
            ElseIf name.Contains("typename") OrElse name.Contains("controltype") Then
                Return "ControlTypes"
            ElseIf name.Contains("fontname") Then
                Return "FontName"
            ElseIf name.Contains("formname") Then
                Return "FormName"
            End If

            Return ""
        End Function

        Private Function GetIdentifier(
                            tokens As List(Of Token),
                           ByRef index As Integer,
                           ByRef line As ITextSnapshotLine
                     ) As Token

            Dim matchingPair1 As Char
            Dim matchingPair2 As Char
            Dim token = tokens(index)

            Select Case token.Type
                Case TokenType.RightParens
                    matchingPair1 = ")"c
                    matchingPair2 = "("c

                Case TokenType.RightBracket
                    matchingPair1 = "]"c
                    matchingPair2 = "["c

                Case TokenType.RightCurlyBracket
                    matchingPair1 = "}"c
                    matchingPair2 = "{"c

                Case Else
                    Return token
            End Select

            Dim pair1Count = 0
            Dim startPos = token.Column + line.Start

            Do
                token = GetNextToken(index, -1, line, tokens)
                Select Case token.Text
                    Case ""
                        Return token
                    Case matchingPair1
                        pair1Count += 1
                    Case matchingPair2
                        If pair1Count = 0 Then
                            Return GetNextToken(index, -1, line, tokens)
                        End If
                        pair1Count -= 1
                End Select
            Loop

            Return token
        End Function


        Dim canGetOriginalBag As Boolean

        Public Function GetOriginalBag() As CompletionBag
            If Not canGetOriginalBag Then Return Nothing

            Dim snapshot = textView.TextSnapshot
            Dim pos = textView.Caret.Position.TextInsertionIndex
            Dim line = snapshot.GetLineFromPosition(pos)
            Dim curToken As Token
            Dim bag = GetCompletionBag(line, pos - line.Start, curToken)

            If bag IsNot Nothing AndAlso bag.CompletionItems.Count > 0 Then
                Return bag
            End If

            Return Nothing
        End Function

        Private Sub ShowCompletionAdornment(
                            snapshot As ITextSnapshot,
                            bag As CompletionBag,
                            curToken As Token,
                            line As ITextSnapshotLine
                     )

            Dim adornmentSpan = GetTextSpanFromToken(line, curToken)
            Dim textSpan = adornmentSpan

            If textSpan.GetSpan(line.TextSnapshot).IsEmpty AndAlso line.TextSnapshot.Length > 0 Then
                If curToken.Column = 0 Then
                    adornmentSpan = New TextSpan(
                        line.TextSnapshot,
                        line.Start,
                        System.Math.Min(curToken.EndColumn + 1, snapshot.Length),
                        SpanTrackingMode.EdgeInclusive
                    )

                Else
                    adornmentSpan = New TextSpan(
                        line.TextSnapshot,
                        line.Start + curToken.Column - 1,
                        System.Math.Min(curToken.EndColumn - curToken.Column + 1, snapshot.Length),
                        SpanTrackingMode.EdgeInclusive
                    )
                End If
            End If

            adornment = New CompletionAdornment(Me, bag, adornmentSpan, textSpan)

            Application.Current.Dispatcher.Invoke(
                DispatcherPriority.Normal, CType(
                Function()
                    RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(adornmentSpan))
                    Return Nothing
                End Function, DispatcherOperationCallback), Nothing)
        End Sub

        Private Function GetCompletionBag(especialItem As String) As CompletionBag
            Dim bag As CompletionBag
            bag = compHelper.GetEmptyCompletionBag(GlobalParser)
            Select Case especialItem
                Case "Boolean"
                    especialItem = "*"
                    bag.CompletionItems.AddRange({
                        New CompletionItem() With {
                            .Key = "False",
                            .DisplayName = "False",
                            .ItemType = CompletionItemType.Keyword
                        },
                        New CompletionItem() With {
                            .Key = "True",
                            .DisplayName = "True",
                            .ItemType = CompletionItemType.Keyword
                        }
                    })

                Case "FontName"
                    bag.CompletionItems.AddRange(FontNames)
                    especialItem = "*"      ' * means that global completion items can appear if user writes names that are not in these especial items.

                Case "FormName"
                    bag.CompletionItems.AddRange(FormNames)
                    especialItem = "*"

                Case "Sub"
                    bag.IsHandler = True
                    CompletionHelper.FillSubroutines(bag)
                    especialItem = "" ' Don't offer any other global items

                Case Else
                    Dim typeInfo As TypeInfo = Nothing
                    If bag.TypeInfoBag.Types.TryGetValue(especialItem.ToLower(), typeInfo) Then
                        CompletionHelper.FillMemberNames(bag, typeInfo, "")
                        bag.CompletionItems.Sort(
                            Function(ci1, ci2)
                                Return ci1.DisplayName.CompareTo(ci2.DisplayName)
                            End Function)
                    End If
            End Select

            If bag IsNot Nothing Then
                bag.SelectEspecialItem = especialItem
            End If

            Return bag
        End Function

        Private Function GetTextSpanFromToken(line As ITextSnapshotLine, token As Token) As ITextSpan
            If token.IsIllegal AndAlso token.Column = 0 Then
                Return line.TextSnapshot.CreateTextSpan(line.End, 0, SpanTrackingMode.EdgeInclusive)
            End If

            Return line.TextSnapshot.CreateTextSpan(
                line.Start + token.Column,
                token.Text.Length,
                SpanTrackingMode.EdgeInclusive
           )

        End Function

        Private Function GetToken(tokenEnumerator As TokenEnumerator, column As Integer) As Token
            Do
                Dim current = tokenEnumerator.Current

                If current.Column <= column AndAlso current.EndColumn > column Then
                    Return current
                End If
            Loop While tokenEnumerator.MoveNext()

            Return Token.Illegal
        End Function

        Private Shared completionItems As New Dictionary(Of String, List(Of CompletionItem))

        Dim WordHighlightColor As Media.Color

        Shared Sub New()
            For Each t In WinForms.PreCompiler.GetTypes()
                AddCompletionList(t)
            Next
        End Sub

        Private Shared Sub AddCompletionList(type As Type)
            Dim compList As New List(Of CompletionItem)

            Dim methods = type.GetMethods(System.Reflection.BindingFlags.Static Or System.Reflection.BindingFlags.Public)

            For Each methodInfo In methods
                Dim name As String
                Dim item As New CompletionItem()
                If methodInfo.GetCustomAttributes(GetType(WinForms.ExMethodAttribute), inherit:=False).Count > 0 Then
                    name = methodInfo.Name
                    item.Key = name
                    item.DisplayName = name
                    item.ItemType = CompletionItemType.MethodName

                    If methodInfo.GetParameters().Length > 1 Then
                        item.ReplacementText = name & "("
                    Else
                        item.ReplacementText = name & "()"
                    End If

                    item.MemberInfo = methodInfo
                    compList.Add(item)

                ElseIf methodInfo.Name.ToLower().StartsWith("get") AndAlso methodInfo.GetCustomAttributes(GetType(WinForms.ExPropertyAttribute), inherit:=False).Count > 0 Then
                    name = methodInfo.Name.Substring(3)
                    item.Key = name
                    item.DisplayName = name
                    item.ItemType = CompletionItemType.PropertyName
                    item.ReplacementText = name
                    item.MemberInfo = methodInfo
                    compList.Add(item)
                End If
            Next

            Dim events = type.GetEvents(System.Reflection.BindingFlags.Static Or System.Reflection.BindingFlags.Public)
            For Each eventInfo In events
                If eventInfo.EventHandlerType Is GetType(Library.SmallVisualBasicCallback) Then
                    Dim name = eventInfo.Name
                    compList.Add(New CompletionItem() With {
                    .Key = name,
                    .DisplayName = name,
                    .ItemType = CompletionItemType.EventName,
                    .ReplacementText = name,
                    .MemberInfo = eventInfo
                })
                End If
            Next

            completionItems.Add(type.Name, compList)
        End Sub

        Private Sub FillMemberNames(
                                completionBag As CompletionBag,
                                moduleName As String,
                                objName As String
                     )

            Dim controlModule = NameOf(WinForms.Control)
            Select Case moduleName
                Case NameOf(WinForms.ImageBox),
                         NameOf(WinForms.Forms),
                         NameOf(WinForms.ColorEx),
                         NameOf(WinForms.TextEx),
                         NameOf(WinForms.ArrayEx),
                         NameOf(WinForms.MathEx),
                         NameOf(WinForms.DateEx),
                         NameOf(WinForms.WinTimer)

                Case Else
                    For Each item In completionItems(controlModule)
                        item.ObjectName = objName
                        completionBag.CompletionItems.Add(item)
                    Next

            End Select

            If moduleName <> controlModule AndAlso completionItems.ContainsKey(moduleName) Then
                For Each item In completionItems(moduleName)
                    item.ObjectName = objName
                    completionBag.CompletionItems.Add(item)
                Next
            End If

        End Sub

        Public Sub ShowHelp(Optional force As Boolean = False)
            OnHelpUpdate(force, Nothing)
        End Sub
        Public Sub ShowHelp(symbol As String)
            OnHelpUpdate((symbol, True), Nothing)
        End Sub
    End Class
End Namespace
