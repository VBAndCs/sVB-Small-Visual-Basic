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
        Private completionHelper As CompletionHelper
        Private adornment As CompletionAdornment
        Private dismissedSpan As ITextSpan
        Private helpUpdateTimer As DispatcherTimer
        Private editorOperations As IEditorOperations
        Private undoHistory As UndoHistory
        Public Event AdornmentsChanged As EventHandler(Of AdornmentsChangedEventArgs) Implements IAdornmentProvider.AdornmentsChanged

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

            If Not textBuffer.Properties.TryGetProperty(GetType(CompletionHelper), completionHelper) Then
                completionHelper = New CompletionHelper()
                textBuffer.Properties.AddProperty(GetType(CompletionHelper), completionHelper)
            End If
        End Sub

        Public Shared CompHistory As New Dictionary(Of String, String)

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

        Private Function GetGlobalParser() As Parser
            If _document Is Nothing Then
                _document = textBuffer.Properties.GetProperty("Document")
            End If

            Dim parsers As List(Of Parser) = _document.CompileGlobalModule()
            If parsers Is Nothing OrElse parsers.Count = 0 Then
                Return Nothing
            Else
                Return parsers(0)
            End If
        End Function


        Public Sub CommitItem(itemWrapper As CompletionItemWrapper)
            Dim item = itemWrapper.CompletionItem
            If adornment IsNot Nothing Then
                Dim key = item.HistoryKey
                If key <> "" Then CompHistory(key) = item.DisplayName
                Dim repText = itemWrapper.ReplacementText
                Dim replaceSpan = GetReplacementSpane()

                If replaceSpan.Length = 0 And replaceSpan.Start > 0 Then
                    Select Case textView.TextSnapshot(replaceSpan.Start - 1)
                        Case "=", "<", ">", "+", "-", "*", "/"
                            repText = " " + repText
                    End Select
                End If

                editorOperations.ReplaceText(replaceSpan, repText, undoHistory)
                DismissAdornment(force:=True)
                TryCast(textView, Control)?.Focus()
                If repText.EndsWith("(") Then ShowHelp()
            End If
        End Sub

        Friend Function GetReplacementSpane() As SnapshotSpan
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

            If TypeOf sender Is Boolean Then
                force = CType(sender, Boolean)
            ElseIf TypeOf sender Is String Then
                symbol = CType(sender, String)
            End If

            Dim span As New Span(pos, 0)

            If Not ParseSepTokens(
                    tokens, startLine, currentLine, endLine,
                    symbol, force, line, column, span,
                    paramIndex) Then Return

            If sourceCodeChanged Then
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
                Return
            End If

            If symbol <> "" AndAlso symbol <> "(" AndAlso line.Start + column = pos Then Return

            ShowHelpInfo(line, column, paramIndex, span)

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

                For i = 0 To n
                    Dim token = tokens(i)
                    If token.Contains(currentLine, column, Not textView.Selection.IsEmpty) Then
                        prevToken = GetNonCommentToken(tokens, i - 1, True)
                        prevIsSep = IsPrevSeparator(currentLine, symbol, prevToken, prevChar)
                        exactToken = token

                        If column > token.Column Then
                            currentToken = If(prevIsSep AndAlso token.ParseType = ParseType.Comment,
                                                        Token.Illegal,
                                                         token)
                            nextToken = GetNonCommentToken(tokens, i + 1, False)

                        ElseIf Not (prevIsSep AndAlso textView.Selection.IsEmpty) Then
                            currentToken = token
                            nextToken = GetNonCommentToken(tokens, i + 1, False)
                        End If
                        Exit For

                    ElseIf token.Line > currentLine Then
                        prevToken = GetNonCommentToken(tokens, i - 1, True)
                        prevIsSep = IsPrevSeparator(currentLine, symbol, prevToken, prevChar)
                        nextToken = GetNonCommentToken(tokens, i, False)
                        Exit For

                    ElseIf token.Line = currentLine Then
                        If token.EndColumn = column Then ' pos is right after the end of the current token
                            exactToken = token
                            prevToken = GetNonCommentToken(tokens, i, True)
                            prevIsSep = IsPrevSeparator(currentLine, symbol, prevToken, prevChar)
                            nextToken = GetNonCommentToken(tokens, i + 1, True)
                            Exit For

                        ElseIf token.Column > column OrElse i = n Then
                            prevToken = GetNonCommentToken(tokens, i - 1, True)
                            prevIsSep = IsPrevSeparator(currentLine, symbol, prevToken, prevChar)
                            Exit For
                        End If
                    End If
                Next

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
                    If currentToken.IsIllegal AndAlso symbol = "" Then
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

            If commaOrParans Then
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

            ElseIf symbol = "(" Then
                paramIndex = 0
            End If

            Return True
        End Function

        Private Sub ShowHelpInfo(
                       line As ITextSnapshotLine,
                       column As Integer,
                       paramIndex As Integer,
                       span As Span
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

                ElseIf item.DisplayName.ToLower() = tokenText Then
                    If bag.IsMethod Then
                        If item.ItemType <> CompletionItemType.MethodName Then Continue For
                    End If

                    Dim spans = HeighLightIdentifiers(currentToken, bag.SubroutineName, item.ItemType, bag.SymbolTable, textView.TextSnapshot)?.ToArray()
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
                If Char.IsLetter(c) Then
                    ShowCompletionAdornment(e.After, newEnd)
                Else
                    Select Case c
                        Case "."c, "!"c
                            popHelp.IsOpen = False
                            ShowCompletionAdornment(e.After, newEnd)

                        Case "="c, ">"c, "<"c
                            ShowCompletionAdornment(e.After, newEnd, True)

                        Case "("c
                            OnHelpUpdate("(", Nothing)
                        Case ")"c
                            OnHelpUpdate(")", Nothing)
                        Case ","c
                            OnHelpUpdate(",", Nothing)
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
            Dim globalParser = GetGlobalParser()
            If needsToReCompile Then
                completionHelper.Compile(source, controlNames, controlsInfo)
                needsToReCompile = False
            End If

            Dim newBag = completionHelper.GetEmptyCompletionBag(globalParser)
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
                        ElseIf method.Contains("showform") OrElse method.Contains("showdialog") Then
                            newBag.CompletionItems.AddRange(FormNames)
                        End If
                    End If
                    Return newBag

                ElseIf controlsInfo IsNot Nothing Then
                    Dim txt = currentToken.LCaseText
                    If txt <> "" AndAlso controlNames IsNot Nothing Then
                        Dim controls = From name In controlNames
                                       Where name(0) <> "("c AndAlso (
                                           (forHelp AndAlso name.ToLower() = txt) OrElse
                                           (Not forHelp AndAlso name.ToLower().StartsWith(txt))
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
                    completionHelper.FillDynamicMembers(newBag, identifierToken.Text)

                ElseIf controlsInfo?.ContainsKey(name) Then
                    FillMemberNames(newBag, controlsInfo(name), identifierToken.Text)

                Else
                    Dim moduleName = completionHelper.GetInferedType(identifierToken)
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

                    Dim bag = completionHelper.GetCompletionItems(
                        line.LineNumber, column,
                        prevToken.Type = TokenType.Equals OrElse currentToken.Type = TokenType.Equals,
                        IsCompletionOperator(prevToken) OrElse IsCompletionOperator(currentToken),
                        forHelp,
                        globalParser
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
                           Optional checkEspecialItem As Boolean = False
                    )

            ' Wait until code editor respond to changes, to avoid any conflicts
            ShowCompletion.After(
                  20,
                  Sub() DoShowCompletionAdornment(snapshot, caretPosition, checkEspecialItem)
            )
        End Sub

        Public Sub DoShowCompletionAdornment(
                           snapshot As ITextSnapshot,
                           caretPosition As Integer,
                           Optional checkEspecialItem As Boolean = False
                    )

            Dim bag As CompletionBag
            Dim especialItem = ""
            Dim leadingSpace = ""
            Dim curToken As Token

            Dim line = snapshot.GetLineFromPosition(caretPosition)
            If checkEspecialItem Then
                Dim tokens = LineScanner.GetTokens(line.GetText(), line.LineNumber)
                If tokens.Count < 2 Then Return

                Dim n = GetLastTokenIndex(line, caretPosition, tokens)
                If n = 0 Then Return

                Select Case tokens(n).Type
                    Case TokenType.Equals, TokenType.NotEqualTo,
                             TokenType.GreaterThan, TokenType.GreaterThanEqualTo,
                             TokenType.LessThan, TokenType.LessThanEqualTo

                        Dim index = n - 1
                        Dim matchingPair1 As Char
                        Dim matchingPair2 As Char
                        Dim token = tokens(index)

                        curToken = tokens(n)
                        If caretPosition - line.Start - tokens(n).EndColumn = 0 Then leadingSpace = " "

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
                                GoTo LineElse
                        End Select

                        Dim pair1Count = 0
                        Dim startPos = token.Column + line.Start

                        Do
                            token = GetNextToken(index, -1, line, tokens)
                            Select Case token.Text
                                Case ""
                                    Return
                                Case matchingPair1
                                    pair1Count += 1
                                Case matchingPair2
                                    If pair1Count = 0 Then
                                        token = GetNextToken(index, -1, line, tokens)
                                        Exit Do
                                    End If
                                    pair1Count -= 1
                            End Select
                        Loop

LineElse:
                        Dim name = tokens(index).LCaseText

                        If name.StartsWith("color") OrElse name.EndsWith("color") OrElse name.EndsWith("colors") Then
                            especialItem = "Colors"
                        ElseIf name.StartsWith("key") OrElse name.EndsWith("key") OrElse name.EndsWith("keys") Then
                            especialItem = "Keys"
                        ElseIf name = "showdialog" OrElse name.StartsWith("dialogresult") OrElse name.EndsWith("dialogresult") Then
                            especialItem = "DialogResults"
                        Else
                            Dim properties = textBuffer.Properties
                            Dim controlsInfo = properties.GetProperty(Of Dictionary(Of String, String))("ControlsInfo")
                            Dim controlNames = properties.GetProperty(Of List(Of String))("ControlNames")
                            Dim source = New TextBufferReader(line.TextSnapshot)
                            GetGlobalParser()
                            Dim symbolTable = completionHelper.Compile(
                                source,
                                controlNames,
                                controlsInfo
                            ).Parser.SymbolTable

                            needsToReCompile = False
                            token.Parent = completionHelper.GetStatement(line.LineNumber)

                            Select Case symbolTable.GetInferedType(token)
                                Case VariableType.Color
                                    especialItem = "Colors"
                                Case VariableType.Key
                                    especialItem = "Keys"
                                Case VariableType.DialogResult
                                    especialItem = "DialogResults"
                                Case Else
                                    Return
                            End Select
                        End If
                End Select
            End If

            If especialItem = "" Then
                bag = GetCompletionBag(line, caretPosition - line.Start, curToken)
            Else
                Dim typeInfo As TypeInfo = Nothing
                bag = completionHelper.GetEmptyCompletionBag(GetGlobalParser())
                If bag.TypeInfoBag.Types.TryGetValue(especialItem.ToLower(), typeInfo) Then
                    CompletionHelper.FillMemberNames(bag, typeInfo, "")
                    bag.CompletionItems.Sort(
                        Function(ci1, ci2) ci1.DisplayName.CompareTo(ci2.DisplayName))
                End If
            End If

            If bag Is Nothing OrElse bag.CompletionItems.Count <= 0 Then
                Return
            End If

            bag.SelectEspecialItem = especialItem

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
                If eventInfo.EventHandlerType Is GetType(Library.SmallBasicCallback) Then
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
                         NameOf(WinForms.DateEx)

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
    End Class
End Namespace
