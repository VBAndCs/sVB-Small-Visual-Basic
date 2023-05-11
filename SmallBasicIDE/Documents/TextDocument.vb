Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Core.Undo.DataExports
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations
Imports Microsoft.SmallVisualBasic.Shell
Imports Microsoft.Windows.Controls
Imports Microsoft.SmallVisualBasic.WinForms
Imports System
Imports System.Collections.ObjectModel
Imports System.ComponentModel.Composition
Imports System.Globalization
Imports System.IO
Imports System.Windows.Media
Imports System.Windows.Threading
Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Linq
Imports Microsoft.SmallVisualBasic.LanguageService
Imports System.Windows
Imports Microsoft.SmallBasic

Namespace Microsoft.SmallVisualBasic.Documents
    Public Class TextDocument
        Inherits FileDocument
        Implements INotifyImport

        Private saveMarker As UndoTransactionMarker
        Private _textBuffer As ITextBuffer
        Private textBufferUndoManager As ITextBufferUndoManager
        Private _undoHistory As UndoHistory
        Private _editorControl As CodeEditorControl
        Private _errorListControl As ErrorListControl
        Private _caretPositionText As String
        Private _programDetails As com.smallbasic.ProgramDetails

        Public LastModified As Date = Date.Now

        Public Property BaseId As String

        Public Property WordHighlightColor As Media.Color = Media.Colors.LightGray

        Public Property MatchingPairsHighlightColor As Media.Color = Media.Colors.LightGreen

        Public Property CaretPositionText As String
            Get
                Return _caretPositionText
            End Get

            Private Set(value As String)
                _caretPositionText = value
                NotifyProperty("CaretPositionText")
            End Set
        End Property

        Public Property MdiView As MdiView

        Public Property ContentType As String
            Get
                Return TextBuffer.ContentType
            End Get

            Set(value As String)
                TextBuffer.ContentType = value
                NotifyProperty("ContentType")
            End Set
        End Property

        Public ReadOnly Property Errors As New ObservableCollection(Of String)

        Public Property ProgramDetails As com.smallbasic.ProgramDetails
            Get
                Return _programDetails
            End Get

            Set(value As com.smallbasic.ProgramDetails)
                _programDetails = value
                NotifyProperty("ProgramDetails")
            End Set
        End Property

        Public ReadOnly Property TextBuffer As ITextBuffer
            Get
                If _textBuffer Is Nothing Then CreateBuffer()
                Return _textBuffer
            End Get
        End Property

        Public ReadOnly Property EditorControl As CodeEditorControl
            Get
                If _editorControl Is Nothing Then
                    _editorControl = New CodeEditorControl With {
                        .TextBuffer = TextBuffer
                    }
                    AddHandler _editorControl.KeyUp, AddressOf AutoCompleteBlocks

                    App.GlobalDomain.AddComponent(_editorControl)
                    App.GlobalDomain.Bind()
                    _editorControl.HighlightSearchHits = True
                    _editorControl.ScaleFactor = 1.0
                    _editorControl.Background = Brushes.Transparent
                    _editorControl.EditorOperations.TabSize = 2
                    _editorControl.IsLineNumberMarginVisible = True
                    _editorControl.Focus()

                    AddHandler _editorControl.TextView.Caret.PositionChanged, AddressOf OnCaretPositionChanged

                    _editorControl.Dispatcher.BeginInvoke(DispatcherPriority.Render,
                          CType(Function()
                                    Me.Focus()
                                    UpdateCaretPositionText()
                                    Return Nothing
                                End Function,
                                DispatcherOperationCallback), Nothing)
                End If

                Return _editorControl
            End Get
        End Property

        Public ReadOnly Property ErrorListControl As ErrorListControl
            Get

                If _errorListControl Is Nothing Then
                    _errorListControl = New ErrorListControl(Me)
                    _errorListControl.ItemsSource = Errors
                End If

                Return _errorListControl
            End Get
        End Property

        Public ReadOnly Property Text As String
            Get
                Dim snapshot = TextBuffer.CurrentSnapshot
                Return snapshot.GetText(0, snapshot.Length)
            End Get
        End Property

        Public Overrides ReadOnly Property Title As String
            Get
                Dim text = If(IsDirty, " *", "")

                If _form <> "" Then Return $"{_form}{text}"

                If IsNew Then
                    If BaseId = "" Then Return ResourceHelper.GetString("Untitled") & text

                    Return String.Format(CultureInfo.CurrentCulture, ResourceHelper.GetString("ImportedProgram"), New Object(0) {BaseId}) & text
                End If

                Return $"{Path.GetFileName(Me.File)}{text}"
            End Get
        End Property

        Public ReadOnly Property UndoHistory As UndoHistory
            Get
                Return _undoHistory
            End Get
        End Property

        <Import>
        Public Property TextBufferFactory As ITextBufferFactory

        <Import>
        Public Property TextBufferUndoManagerProvider As ITextBufferUndoManagerProvider

        <Import>
        Public Property UndoHistoryRegistry As IUndoHistoryRegistry

        Public GlobalModuleHasChanged As Boolean

        Public Sub New(filePath As String)
            MyBase.New(filePath)
            saveMarker = New UndoTransactionMarker()
            App.GlobalDomain.AddComponent(Me)
            App.GlobalDomain.Bind()
            _ControlNames.Add("(Global)")
            _GlobalSubs.Add(AddNewFunc)
            _GlobalSubs.Add(AddNewSub)
            _ControlEvents.Add(AddNewFunc)
            _ControlEvents.Add(AddNewSub)
            ParseFormHints()
            UpdateGlobalSubsList()
            AddProperty("Document", Me)
        End Sub


        Private Sub OnCaretPositionChanged(sender As Object, e As CaretPositionChangedEventArgs)
            stopFormatingLine = -1
            If IgnoreCaretPosChange OrElse StillWorking Then Return
            If _formatting Then
                _formatting = False
                Return
            End If

            EditorControl.Dispatcher.BeginInvoke(
                    Sub()
                        StillWorking = True
                        If needsToFormat And e IsNot Nothing Then
                            Dim snapshot = e.TextView.TextSnapshot
                            Dim pos = _editorControl.TextView.Caret.Position.TextInsertionIndex

                            Dim pos1 = e.OldPosition.TextInsertionIndex
                            If pos1 > snapshot.Length Then pos1 = pos

                            Dim pos2 = e.NewPosition.TextInsertionIndex
                            If pos2 > snapshot.Length Then pos2 = pos

                            If pos1 <> pos2 Then
                                Dim line1 = snapshot.GetLineNumberFromPosition(pos1)
                                Dim line2 = snapshot.GetLineNumberFromPosition(pos2)
                                If line1 <> line2 Then
                                    needsToFormat = False
                                    FormatLine(line1)
                                End If
                            End If
                        End If

                        Try
                            Dim handlerName = GetCurrentSubName()
                            _MdiView.FreezeCmbEvents = True
                            If _EventHandlers.ContainsKey(handlerName) Then
                                UpdateCombos(_EventHandlers(handlerName))

                            Else
                                Dim eventInfo = GetHandlerInfo(handlerName)
                                If eventInfo.ControlName <> "" Then
                                    ' Restore a broken handler. This can happen when deleting a control then restoring it.
                                    _EventHandlers.Add(handlerName, eventInfo)
                                    UpdateCombos(eventInfo)

                                ElseIf _MdiView.CmbControlNames IsNot Nothing Then ' Global
                                    _MdiView.CmbControlNames.SelectedIndex = 0
                                    If handlerName = "" Then
                                        _MdiView.SelectEventName("-1") ' Select item at index -1
                                    Else
                                        _MdiView.SelectEventName(handlerName)
                                    End If
                                End If
                            End If
                            _MdiView.FreezeCmbEvents = False

                            If sourceCodeChanged OrElse Not (
                                            _editorControl.ContainsWordHighlights AndAlso
                                            _editorControl.IsHighlighted(_editorControl.TextView.Caret.Position.TextInsertionIndex)
                                        ) Then

                                If sourceCodeChanged Then
                                    sourceCodeChanged = False
                                    needsToreCompile = True
                                    _editorControl.ClearHighlighting()
                                End If

                                HighlightMatchingPair()
                            End If

                            UpdateCaretPositionText()

                        Finally
                            StillWorking = False
                        End Try
                    End Sub,
                 DispatcherPriority.ContextIdle)

        End Sub

        Sub HighlightMatchingPair()
            Dim textView = EditorControl.TextView
            Dim snapshot = textView.TextSnapshot
            Dim length = snapshot.Length

            If length = 0 Then Return

            Dim sel = textView.Selection
            If sel.ActiveSnapshotSpan.Length > 1 Then
                Dim span = sel.ActiveSnapshotSpan
                Dim tokenType = LineScanner.GetTokenType(span.GetText().ToLower())
                Dim parseType = LineScanner.GetParseType(tokenType)

                If parseType = ParseType.Keyword Then
                    HighlightBlockKeywords(snapshot.GetLineFromPosition(span.Start).LineNumber, snapshot, tokenType)
                End If

                Return
            End If

            Dim pos = If(sel.IsEmpty,
                    textView.Caret.Position.TextInsertionIndex,
                    sel.ActiveSnapshotSpan.Start
            )

            Dim line = snapshot.GetLineFromPosition(pos)
            Dim tokens = LineScanner.GetTokens(line.GetText(), 0)
            Dim index = GetTokenAt(pos - line.Start, tokens)
            If index = -1 Then Return

            Dim token = tokens(index)
            Dim matchingPair1 As Char
            Dim matchingPair2 As Char
            Dim direction As Integer = 1

            Select Case token.Type
                Case TokenType.LeftParens
                    matchingPair1 = "("c
                    matchingPair2 = ")"c

                Case TokenType.LeftBracket
                    matchingPair1 = "["c
                    matchingPair2 = "]"c

                Case TokenType.LeftCurlyBracket
                    matchingPair1 = "{"c
                    matchingPair2 = "}"c

                Case TokenType.RightParens
                    matchingPair1 = ")"c
                    matchingPair2 = "("c
                    direction = -1

                Case TokenType.RightBracket
                    matchingPair1 = "]"c
                    matchingPair2 = "["c
                    direction = -1

                Case TokenType.RightCurlyBracket
                    matchingPair1 = "}"c
                    matchingPair2 = "{"c
                    direction = -1

                Case Else
                    If sel.IsEmpty AndAlso token.ParseType = ParseType.Keyword Then HighlightBlockKeywords(line.LineNumber, snapshot, token.Type)
                    Return
            End Select

            Dim pair1Count = 0
            Dim startPos = token.Column + line.Start

            Do
                token = GetNextToken(index, direction, line, tokens)
                Select Case token.Text
                    Case ""
                        Return

                    Case matchingPair1
                        pair1Count += 1

                    Case matchingPair2
                        If pair1Count = 0 Then
                            If direction < 0 Then
                                _editorControl.HighlightWords(_MatchingPairsHighlightColor, (token.Column + line.Start, 1), (startPos, 1))
                            Else
                                _editorControl.HighlightWords(_MatchingPairsHighlightColor, (startPos, 1), (token.Column + line.Start, 1))
                            End If
                            Return
                        End If
                        pair1Count -= 1
                End Select
            Loop

        End Sub


        Private Function GetTokenAt(index As Integer, tokens As List(Of Token)) As Integer
            For i = 0 To tokens.Count - 1
                Dim token = tokens(i)
                If index <= token.EndColumn Then
                    If index < token.Column Then Return -1
                    Return i
                End If
            Next
            Return -1
        End Function

        Public Sub HighlightEnclosingBlockKeywords()
            Dim textView = _editorControl.TextView
            Dim pos = textView.Caret.Position.TextInsertionIndex
            Dim lineNumber = textView.TextSnapshot.GetLineNumberFromPosition(pos)
            Dim statement = GetBlock(lineNumber, Nothing)
            _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))
        End Sub


        Dim highlightCompiler As New Compiler()

        Private Sub HighlightBlockKeywords(lineNumber As Integer, snapshot As ITextSnapshot, token As TokenType)
            Select Case token
                Case TokenType.Sub, TokenType.EndSub, TokenType.Function, TokenType.EndFunction, TokenType.Return
                    Dim statement = GetBlock(lineNumber, GetType(Statements.SubroutineStatement))
                    _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))

                Case TokenType.If, TokenType.Then, TokenType.ElseIf, TokenType.Else, TokenType.EndIf
                    Dim statement = GetBlock(lineNumber, GetType(Statements.IfStatement))
                    _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))

                Case TokenType.For, TokenType.To, TokenType.Step, TokenType.EndFor
                    Dim statement = GetBlock(lineNumber, GetType(Statements.ForStatement))
                    _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))

                Case TokenType.ForEach, TokenType.In
                    Dim statement = GetBlock(lineNumber, GetType(Statements.ForEachStatement))
                    _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))

                Case TokenType.Next
                    Dim statement1 = GetBlock(lineNumber, GetType(Statements.ForStatement))
                    Dim statement2 = GetBlock(lineNumber, GetType(Statements.ForEachStatement))
                    Dim statement As Statements.Statement
                    If statement1 Is Nothing Then
                        statement = statement2
                    ElseIf statement2 Is Nothing Then
                        statement = statement1
                    ElseIf statement1.StartToken.IsBefore(statement2.StartToken.Line, statement2.StartToken.Column) Then
                        statement = statement2
                    Else
                        statement = statement1
                    End If
                    _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))

                Case TokenType.While, TokenType.Wend, TokenType.EndWhile
                    Dim statement = GetBlock(lineNumber, GetType(Statements.WhileStatement))
                    _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))

                Case TokenType.ContinueLoop, TokenType.ExitLoop
                    Dim statement = GetLoopBlock(lineNumber)
                    _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))
            End Select
        End Sub

        Public Function GetBlock(lineNumber As Integer, statementType As Type) As Statements.Statement
            CompileForHighlight()

            Dim statement = Completion.CompletionHelper.GetStatement(highlightCompiler, lineNumber)

            If statement Is Nothing Then Return Nothing
            statement = statement.GetStatementAt(lineNumber)

            ' Get parent block
            If statement IsNot Nothing Then
                Do
                    If statement.IsOfType(statementType) Then Return statement
                    statement = statement.Parent
                    If statement Is Nothing Then Exit Do
                Loop
            End If

            Return Nothing
        End Function

        Private Sub CompileForHighlight()
            If GetProperty(Of Boolean)("GlobalChanged") Then needsToreCompile = True
            CompileGlobalModule()
            If needsToreCompile Then
                needsToreCompile = False
                highlightCompiler.Compile(New StringReader(Me.Text), True)
            End If
        End Sub

        Public Function GetLoopBlock(lineNumber As Integer) As Statements.Statement
            CompileForHighlight()
            Dim statement = Completion.CompletionHelper.GetStatement(highlightCompiler, lineNumber)
            If statement Is Nothing Then Return Nothing

            Dim jumpStatement = TryCast(statement.GetStatementAt(lineNumber), Statements.JumpLoopStatement)
            If jumpStatement Is Nothing Then Return Nothing

            Dim loopsCount = jumpStatement.UpLevel + 1
            Dim loops = jumpStatement.GetParentLoops()
            If loops.Count = 0 Then Return Nothing
            Return loops(loops.Count - 1)
        End Function

        Private Function GetCurrentSubText(lineNumber As Integer) As String
            Dim snapshot = _editorControl.TextView.TextSnapshot
            Dim start = snapshot.GetLineFromLineNumber(FindCurrentSubStart(snapshot, lineNumber)).Start
            Dim [end] = snapshot.GetLineFromLineNumber(FindCurrentSubEnd(snapshot, lineNumber)).End
            Return Me.Text.Substring(start, [end] - start + 1)
        End Function

        Public Function GetSpans(tokens As List(Of Token)) As Span()
            If tokens Is Nothing Then Return Nothing

            Dim spans As New List(Of Span)
            Dim snapshot = _editorControl.TextView.TextSnapshot
            Dim lineNum = -1
            Dim linestart = 0

            For Each token In tokens
                If lineNum <> token.Line Then
                    lineNum = token.Line
                    linestart = snapshot.GetLineFromLineNumber(token.Line).Start
                End If

                spans.Add(New Span(
                    linestart + token.Column,
                    token.EndColumn - token.Column
                ))
            Next

            Return spans.ToArray()
        End Function


        Sub UpdateCombos(eventInfo As EventInformation)
            If CStr(_MdiView.CmbControlNames.SelectedItem) <> eventInfo.ControlName Then
                _MdiView.CmbControlNames.SelectedItem = eventInfo.ControlName
            End If

            _MdiView.SelectEventName(eventInfo.EventName)
        End Sub

        Public Overrides Sub Close()
            RemoveHandler _editorControl.TextView.Caret.PositionChanged, AddressOf OnCaretPositionChanged
            RemoveHandler _textBuffer.Changed, AddressOf OnTextBufferChanged
            RemoveHandler _undoHistory.UndoRedoHappened, AddressOf UndoRedoHappened
            MyBase.Close()
        End Sub

        Public Sub Save()
            Using stream As Stream = Open()
                stream.SetLength(0L)

                Using writer As New StreamWriter(stream)
                    TextBuffer.CurrentSnapshot.Write(writer)
                End Using
            End Using

            _undoHistory.ReplaceMarkerOnTop(saveMarker, True)
            IsNew = False
            IsDirty = False
            NotifyProperty("Title")
        End Sub

        Public Sub SaveAs(filePath As String)
            Dim filePath2 = Me.File
            Me.File = filePath

            Try
                Save()
            Catch
                Me.File = filePath2
                Throw
            End Try
        End Sub

        Private Sub CreateBuffer()
            If IsNew OrElse Me.File = "" Then
                _textBuffer = New BufferFactory().CreateTextBuffer()

            Else
                Using reader As StreamReader = New StreamReader(Me.File)
                    _textBuffer = New BufferFactory().CreateTextBuffer(reader, GetContentTypeFromFileExtension())
                End Using
            End If

            AddHandler _textBuffer.Changed, AddressOf OnTextBufferChanged
            AddProperty("WordHighlightColor", _WordHighlightColor)
        End Sub

        Private Function GetContentTypeFromFileExtension() As String
            Dim result = "text"

            Select Case Path.GetExtension(Me.File)
                Case ".sb", ".smallbasic"
                    result = "text.smallbasic"

                Case ".cs"
                    result = "text.csharp"

                Case ".xaml", ".xcml"
                    result = "text.xml"

                Case ".xml"
                    result = "text.xml"

                Case ".vb"
                    result = "text.vb"

            End Select

            Return result
        End Function

        Private Sub OnBindCompleted()
            If Not UndoHistoryRegistry.TryGetHistory(TextBuffer, _undoHistory) Then
                _undoHistory = UndoHistoryRegistry.RegisterHistory(TextBuffer)
            End If

            _undoHistory.ReplaceMarkerOnTop(saveMarker, True)
            AddHandler _undoHistory.UndoRedoHappened, AddressOf UndoRedoHappened
            textBufferUndoManager = TextBufferUndoManagerProvider.GetTextBufferUndoManager(TextBuffer)
        End Sub

        Dim _formatting As Boolean

        Dim sourceCodeChanged As Boolean = True
        Dim needsToreCompile As Boolean
        Dim needsToFormat As Boolean = True

        Private Sub OnTextBufferChanged(sender As Object, e As TextChangedEventArgs)
            sourceCodeChanged = True
            needsToFormat = True
            LastModified = Now
            IsDirty = True
            If IsTheGlobalFile Then
                For Each view As MdiView In _MdiView.MdiViews.Items
                    view.Document.GlobalModuleHasChanged = True
                Next
            End If

            _editorControl.ClearHighlighting()
            ErrorListControl.Close()
            If _formatting OrElse _IgnoreCaretPosChange OrElse StillWorking Then Return

            StillWorking = True
            Try
                EditorControl.Dispatcher.BeginInvoke(
                  Sub()
                      StillWorking = True

                      UpdateGlobalSubsList()
                      ' re-format if lines changed
                      For Each change In e.Changes
                          Dim n = CountLines(change.NewText)
                          If change.Delta > 2 AndAlso change.NewText.Trim() <> "" Then
                              If n = 1 Then
                                  FormatSub()
                                  Exit For
                              ElseIf n > 1 Then
                                  FormatDocument(_editorControl.TextView.TextBuffer)
                                  Exit For
                              End If
                          End If

                          n = CountLines(change.OldText)
                          If n = 1 Then
                              FormatSub()
                              Exit For
                          ElseIf n > 1 Then
                              FormatDocument(_editorControl.TextView.TextBuffer)
                              Exit For
                          End If
                      Next

                      StillWorking = False

                      OnCaretPositionChanged(Nothing, Nothing)
                  End Sub, DispatcherPriority.ContextIdle)

            Catch
                StillWorking = False
            End Try

        End Sub

        Private Sub AutoCompleteBlocks(sender As Object, e As System.Windows.Input.KeyEventArgs)
            If (e.KeyboardDevice.Modifiers And Input.ModifierKeys.Control) <> 0 Then Return

            Dim textView = EditorControl.TextView
            Dim text = textView.TextSnapshot
            Dim insertionIndex = textView.Caret.Position.TextInsertionIndex
            If insertionIndex = 0 Then Return

            Dim c = text.GetText(insertionIndex - 1, 1)
            If Char.IsLetterOrDigit(c) OrElse c = "_" Then Return
            Dim line = text.GetLineFromPosition(insertionIndex)
            Dim code = line.GetText()

            Select Case e.Key
                Case System.Windows.Input.Key.Enter, System.Windows.Input.Key.Space, System.Windows.Input.Key.Tab
                    ' Commit chars

                Case System.Windows.Input.Key.Up, System.Windows.Input.Key.Down, System.Windows.Input.Key.PageUp, System.Windows.Input.Key.PageDown
                    If code.Trim(" "c, vbTab) = "" Then
                        textView.Caret.MoveTo(line.End)
                    End If
                    Return

                Case Else
                    If c <> "("c Then Return

            End Select

            Try
                If code.Trim = "" Then
                    If e.Key = System.Windows.Input.Key.Enter Then
                        line = text.GetLineFromLineNumber(line.LineNumber - 1)
                        code = line.GetText()
                    Else
                        Return
                    End If
                End If

                Dim keyword = code.Trim(" "c, vbTab(0), "("c).ToLower()
                Dim paran = If(c = "(", "(", "")
                e.Handled = True

                Select Case keyword
                    Case ""
                        Return

                    Case "if"
                        AutoCompleteBlock(textView, line, code, keyword, $"If {paran} Then#   ", "EndIf", paran.Length)

                    Case "elseif"
                        AutoCompleteBlock(textView, line, code, keyword, $"ElseIf {paran} Then#   ", "", paran.Length)

                    Case "for"
                        AutoCompleteBlock(textView, line, code, keyword, $"For  = ? To ?#   ", "Next", paran.Length)

                    Case "foreach"
                        AutoCompleteBlock(textView, line, code, keyword, $"ForEach  In ?#   ", "Next", paran.Length)

                    Case "while"
                        AutoCompleteBlock(textView, line, code, keyword, $"While {paran}#   ", "Wend", paran.Length)

                    Case "sub"
                        AutoCompleteBlock(textView, line, code, keyword, $"Sub #   ", "EndSub", paran.Length)

                    Case "function"
                        AutoCompleteBlock(textView, line, code, keyword, $"Function #   ", "EndFunction", paran.Length)

                    Case Else
                        e.Handled = False

                End Select

            Catch
            End Try

        End Sub

        Dim stopFormatingLine As Integer = -1

        Private Sub AutoCompleteBlock(
                             textView As IAvalonTextView,
                             line As ITextSnapshotLine,
                             code As String,
                             keyword As String,
                             block As String,
                             endBlock As String,
                             n As Integer
                   )

            Dim start = code.IndexOf(keyword, StringComparison.InvariantCultureIgnoreCase)
            Dim leadingSpace = code.Substring(0, start)
            Dim L = line.Length
            Dim text = textView.TextSnapshot
            Dim addBlockEnd = True
            Dim parser As New Parser
            Dim isSubroutine = (keyword = "sub" OrElse keyword = "function")

            If endBlock <> "" Then
                For i = line.LineNumber + 1 To text.LineCount - 1
                    Dim nextLine = text.GetLineFromLineNumber(i)
                    Dim LineCode = nextLine.GetText()
                    If LineCode.Trim(" "c, vbTab, "("c).ToLower() <> "" Then
                        Dim indent_Keyword = leadingSpace & keyword
                        If (isSubroutine AndAlso (LineCode = "EndSub" OrElse LineCode = "EndFunction")) OrElse LineCode = leadingSpace & endBlock Then
                            L += nextLine.Length + 2
                            Try
                                nextLine = text.GetLineFromLineNumber(i + 1)
                                LineCode = nextLine.GetText().ToLower()
                                If LineCode.Trim(" "c, vbTab, "("c) = "" Then
                                    L += 2
                                End If
                            Catch ex As Exception
                            End Try

                        ElseIf (isSubroutine AndAlso (LineCode.StartsWith("sub") OrElse LineCode.StartsWith("function"))
                                    ) OrElse LineCode.StartsWith(indent_Keyword) Then
                            addBlockEnd = isSubroutine
                            Exit For

                        Else
                            For j = i + 1 To text.LineCount - 1
                                nextLine = text.GetLineFromLineNumber(j)
                                LineCode = nextLine.GetText()

                                If isSubroutine Then
                                    If LineCode.StartsWith("Sub ") OrElse LineCode.StartsWith("Function ") Then
                                        addBlockEnd = True
                                        Exit For
                                    ElseIf LineCode = "EndSub" OrElse LineCode = "EndFunction" Then
                                        addBlockEnd = False
                                        Exit For
                                    End If

                                ElseIf LineCode = leadingSpace & endBlock Then
                                    addBlockEnd = False
                                    Exit For
                                End If
                            Next
                        End If
                        Exit For
                    End If
                    L += nextLine.Length + 2
                Next
            End If

            Dim indent = leadingSpace.Length
            Dim nl = $"{vbCrLf}{leadingSpace}"
            addBlockEnd = addBlockEnd AndAlso endBlock <> ""
            _formatting = True
            EditorControl.EditorOperations.ReplaceText(
                 New Span(line.Start, L),
                leadingSpace &
                    block.Replace("#", If(addBlockEnd, nl, "")) &
                    If(addBlockEnd, nl & endBlock & nl, ""),
                _undoHistory
            )

            textView.Caret.MoveTo(line.Start + indent + Len(keyword) + 1 + n)
            stopFormatingLine = line.LineNumber
            '_formatting = False
        End Sub

        Friend Sub Focus(Optional moveToStart As Boolean = False)
            Dim txtView = CType(EditorControl.TextView, AvalonTextView)
            txtView.VisualElement.Focus()
            If moveToStart Then
                txtView.Caret.MoveTo(0)
            End If
            txtView.Caret.EnsureVisible()
        End Sub

        Private Sub UndoRedoHappened(sender As Object, e As UndoRedoEventArgs)
            Dim __ As Object = Nothing

            If _undoHistory.TryFindMarkerOnTop(saveMarker, __) Then
                IsDirty = False
            Else
                IsDirty = True
            End If
        End Sub

        Dim StillWorking As Boolean = False
        Friend PageKey As String

        Private Sub UpdateCaretPositionText()
            EditorControl.Dispatcher.BeginInvoke(
                 Sub()
                     Dim textView = _editorControl.TextView
                     Dim textInsertionIndex = textView.Caret.Position.TextInsertionIndex
                     Dim line = textView.TextSnapshot.GetLineFromPosition(textInsertionIndex)
                     CaretPositionText = $"{line.LineNumber + 1},{textInsertionIndex - line.Start + 1}"
                 End Sub,
                 DispatcherPriority.ContextIdle)
        End Sub

        Public Sub ImportCompleted() Implements INotifyImport.ImportCompleted
            OnBindCompleted()
        End Sub

        Public ReadOnly Property GlobalSubs As New List(Of String)
        Public ReadOnly Property ControlEvents As New ObservableCollection(Of String)
        Public ReadOnly Property ControlNames As New ObservableCollection(Of String)
        Public Property EventHandlers As New Dictionary(Of String, EventInformation)

        Dim _form As String

        Public Property Form As String
            Get
                Return _form
            End Get

            Set(value As String)
                _form = value
                AddProperty("FormName", value)
                NotifyProperty("Title")
            End Set
        End Property

        Public Property ControlsInfo As Dictionary(Of String, String)


        Function GenerateCodeBehind(
                             formDesigner As DiagramHelper.Designer,
                             updateControlInfo As Boolean
                        ) As String

            If _form = "" Then Return ""

            Dim genCode As New Text.StringBuilder
            Dim declaration As New Text.StringBuilder
            Dim formName = _form
            genCode.AppendLine("'@Form Hints:")
            genCode.AppendLine($"'#{formName}{{")
            Dim controlsInfoList As New Dictionary(Of String, String)
            Dim controlNamesList As New List(Of String)
            controlNamesList.Add("(Global)")

            controlsInfoList(formName.ToLower()) = "Form"
            controlsInfoList("me") = "Form"
            controlNamesList.Add(formName)
            declaration.AppendLine($"Me = ""{formName.ToLower()}""")

            For Each c As UIElement In formDesigner.Items
                Dim name = formDesigner.GetControlNameOrDefault(c)

                If name <> "" Then
                    Dim typeName = PreCompiler.GetModuleName(c.GetType().Name)
                    controlsInfoList(name.ToLower()) = typeName
                    controlNamesList.Add(name)
                    genCode.AppendLine($"'    {name}: {typeName}")
                    declaration.AppendLine($"{name} = ""{formName.ToLower()}.{name.ToLower()}""")
                End If
            Next

            genCode.AppendLine("'}")
            genCode.AppendLine()
            genCode.Append(declaration)
            ' Take the xaml path if exists, to consider the file name change in save as case.
            Dim xamlFile = If(formDesigner.XamlFile = "",
                    Path.GetFileNameWithoutExtension(Me.File) & ".xaml",
                    IO.Path.GetFileName(formDesigner.XamlFile)
            )
            genCode.AppendLine($"_path = Program.Directory + ""\{xamlFile}""")
            genCode.AppendLine($"{formName} = Forms.LoadForm(""{formName}"", _path)")

            If formDesigner.AllowTransparency Then
                genCode.AppendLine($"Form.AllowTransparency({formName})")
            End If

            genCode.AppendLine($"Form.SetArgsArr({formName.ToLower}, Stack.PopValue(""_{formName.ToLower()}_argsArr""))")
            genCode.AppendLine($"Form.Show({formName})")

            ' Remove Hamdlers of deleted or renamed controls
            For i = _EventHandlers.Count - 1 To 0 Step -1
                If Not controlNamesList.Contains(_EventHandlers.Values(i).ControlName) Then
                    _EventHandlers.Remove(_EventHandlers.Keys(i))
                End If
            Next

            If updateControlInfo Then
                _form = formName
                _ControlsInfo = controlsInfoList
                AddProperty("ControlsInfo", _ControlsInfo)
                AddProperty("ControlNames", controlNamesList)

                ' Note that ControlNames Property is bound to a combobox, so keep the existing collection
                controlNamesList.Sort()
                _ControlNames.Clear()
                If controlNamesList.Count > 0 Then
                    For Each controlName In controlNamesList
                        _ControlNames.Add(controlName)
                    Next

                    If MdiView?.CmbControlNames IsNot Nothing Then
                        MdiView.CmbControlNames.SelectedIndex = 0
                    End If
                End If
            End If

            GenerateEventHints(genCode)

            Return genCode.ToString()
        End Function

        Friend Sub GenerateEventHints(genCode As Text.StringBuilder)

            ' Remove handlers of renamed subs
            RemoveBrokenHandlers()

            If EventHandlers.Count > 0 Then
                genCode.AppendLine("'#Events{")
                Dim sbHandlers As New Text.StringBuilder

                Dim controlEvents = From eventHandler In EventHandlers
                                    Group By eventHandler.Value.ControlName
                                    Into EventInfo = Group

                For Each ev In controlEvents
                    genCode.Append($"'    {ev.ControlName}:")
                    sbHandlers.AppendLine($"' {ev.ControlName} Events:")
                    sbHandlers.AppendLine($"Control.HandleEvents({ev.ControlName})")

                    For Each info In ev.EventInfo
                        Dim controlName = ev.ControlName.ToLower
                        If _ControlsInfo.ContainsKey(controlName) Then
                            genCode.Append($" {info.Value.EventName}")
                            Dim module1 = _ControlsInfo(controlName)
                            Dim moduleName = PreCompiler.GetEventModule(module1, info.Value.EventName)
                            sbHandlers.AppendLine($"{moduleName}.{info.Value.EventName} = {info.Key}")
                        End If
                    Next
                    genCode.AppendLine()
                    sbHandlers.AppendLine()
                Next

                genCode.AppendLine("'}")
                genCode.AppendLine()
                genCode.AppendLine(sbHandlers.ToString())
            End If
        End Sub

        Sub AddProperty(name As String, value As Object)
            TextBuffer.Properties.AddOrModifyProperty(name, value)
        End Sub

        Function GetProperty(Of T)(name As String) As T
            Dim value As T = Nothing
            TextBuffer.Properties.TryGetProperty(name, value)
            Return value
        End Function

        Public Function GetCodeBehind(Optional ToCompile As Boolean = False) As String
            Dim codeFileHasHints = Me.Text.Contains("'@Form Hints:")

            Dim genCodefile = If(File = "", "", Me.File.Substring(0, Me.File.Length - 2) + "sb.gen")
            If genCodefile = "" OrElse Not IO.File.Exists(genCodefile) Then
                Return If(ToCompile, "", Me.Text)
            Else
                If ToCompile Then
                    Return IO.File.ReadAllText(genCodefile)
                Else
                    Return If(codeFileHasHints, Me.Text, IO.File.ReadAllText(genCodefile))
                End If
            End If

        End Function

        Public Sub ParseFormHints()
            Dim code = GetCodeBehind()

            Dim info = PreCompiler.ParseFormHints(code)
            If info Is Nothing Then Return

            Me.Form = info.Form
            Me.ControlsInfo = info.ControlsInfo

            If info.EventHandlers IsNot Nothing Then
                _EventHandlers = info.EventHandlers
            Else
                _EventHandlers.Clear()
            End If

            info.ControlNames.Sort()
            _ControlNames.Clear()
            _ControlNames.Add("(Global)")

            For Each c In info.ControlNames
                _ControlNames.Add(c)
            Next

            AddProperty("ControlsInfo", _ControlsInfo)
            AddProperty("ControlNames", info.ControlNames)
        End Sub

        Friend Shared Function FromCode(code As String) As TextDocument
            Dim n = New Random().Next(1, 1000000)
            Dim filename = Path.Combine(Path.GetTempPath(), $"file{n}.sb")
            Global.My.Computer.FileSystem.WriteAllText(filename, code, False)
            Return New TextDocument(filename)
        End Function

        Private Const SubBodyLength As Integer = 19
        Private Const AddNewSub As String = "(Add New Sub)"
        Private ReadOnly eventHandlerSub As String = vbCrLf & "
'------------------------------------------------
Sub #( )
   
EndSub
"

        Private Const FuncBodyLength As Integer = 37
        Private Const AddNewFunc As String = "(Add New Function)"
        Private ReadOnly globalFunc As String = vbCrLf & "
'------------------------------------------------
Function #( )
   
   Return 0
EndFunction
"
        Public Property IgnoreCaretPosChange As Boolean

        Public Function AddEventHandler(
                           controlName As String,
                           eventName As String,
                           Optional selectSubName As Boolean = True
                    ) As Boolean

            Dim alreadyExists = False
            Dim isGlobal = (controlName = "(Global)")
            Dim handlerName = If(isGlobal, "", controlName & "_") &
                        If(eventName.StartsWith("(Add New "), "", eventName)

            Dim pos = -1
            If handlerName = "" Then
                pos = -1

            ElseIf isGlobal AndAlso _GlobalSubs.Contains(handlerName) Then
                pos = FindEventHandler(handlerName)

            ElseIf _EventHandlers.ContainsKey(handlerName) Then
                pos = FindEventHandler(handlerName)

            Else ' Restore Broken Handler
                pos = FindEventHandler(handlerName)
                If pos > -1 Then
                    _EventHandlers(handlerName) = New EventInformation(controlName, eventName)
                End If
            End If

            Dim caret = EditorControl.TextView.Caret

            _IgnoreCaretPosChange = True
            _editorControl.EditorOperations.ResetSelection()

            If pos = -1 Then
                caret.MoveTo(Me.Text.Length)
                If handlerName = "" Then
                    Dim isSub = (eventName = AddNewSub)
                    Dim NewName = If(isSub, "NewSub_", "NewFunc_")
                    Dim n = 0
                    Dim L = NewName.Length
                    Try
                        n = Aggregate s In _GlobalSubs
                                 Where s.StartsWith(NewName)
                                 Let x = s.Substring(L)
                                 Into Max(If(IsNumeric(x), CInt(x), 0))
                    Catch
                    End Try

                    handlerName = NewName & n + 1
                    Dim name = If(isSub, eventHandlerSub, globalFunc)
                    Dim handler = name.Replace("#", handlerName)
                    _GlobalSubs.Add(handlerName)
                    _editorControl.EditorOperations.InsertText(handler, _undoHistory)
                    caret.MoveTo(Text.Length - If(isSub, SubBodyLength, FuncBodyLength))
                    If selectSubName Then SelectCurrentWord()
                    _ControlEvents.Add(handlerName)

                Else
                    _EventHandlers(handlerName) = New EventInformation(controlName, eventName)
                    Dim handler = eventHandlerSub.Replace("#", handlerName)
                    _editorControl.EditorOperations.InsertText(handler, _undoHistory)
                    EnsureAtTop(Text.Length - 10)
                    alreadyExists = True
                End If

            Else
                alreadyExists = True
                caret.MoveTo(pos)
                SelectCurrentWord()
            End If

            Me.Focus()

            MdiView.SelectEventName(controlName, If(isGlobal, handlerName, eventName))

            If Not (selectSubName OrElse alreadyExists) Then
                _editorControl.EditorOperations.MoveLineDown(False)
                _editorControl.EditorOperations.MoveToEndOfLine(False)
            End If

            _IgnoreCaretPosChange = False
            Return Not alreadyExists
        End Function

        Public Sub SelectWordAt(
                        line As Integer,
                        column As Integer,
                        Optional viewAtTop As Boolean = True
                   )

            If line < 0 Then Return

            Dim currentSnapshot = _textBuffer.CurrentSnapshot
            If line < currentSnapshot.LineCount Then
                Dim lineSnap = currentSnapshot.GetLineFromLineNumber(line)

                If column < lineSnap.LengthIncludingLineBreak Then
                    Dim charIndex = lineSnap.Start + column
                    _editorControl.EditorOperations.ResetSelection()
                    _editorControl.TextView.Caret.MoveTo(charIndex)
                    SelectCurrentWord(viewAtTop)
                End If
            End If
        End Sub

        Public Sub SelectCurrentWord(Optional viewAtTop As Boolean = True)
            Dim caret = _editorControl.TextView.Caret
            Dim ops = _editorControl.EditorOperations
            Dim sv = _editorControl.TextView.ViewScroller

            ops.ResetSelection()
            caret.EnsureVisible()
            sv.ScrollViewportVerticallyByPage(Nautilus.Text.Editor.ScrollDirection.Down)
            caret.EnsureVisible()
            sv.ScrollViewportVerticallyByLine(Nautilus.Text.Editor.ScrollDirection.Up)

            If Not viewAtTop Then
                sv.ScrollViewportVerticallyByLine(Nautilus.Text.Editor.ScrollDirection.Up)
                sv.ScrollViewportVerticallyByLine(Nautilus.Text.Editor.ScrollDirection.Up)
                sv.ScrollViewportVerticallyByLine(Nautilus.Text.Editor.ScrollDirection.Up)
            End If

            ops.SelectCurrentWord()
            Focus()
        End Sub

        Sub EnsureAtTop(pos As Integer)
            Dim caret = _editorControl.TextView.Caret
            Dim ops = _editorControl.EditorOperations
            Dim sv = _editorControl.TextView.ViewScroller

            ops.ResetSelection()
            caret.MoveTo(pos)
            sv.ScrollViewportVerticallyByPage(Nautilus.Text.Editor.ScrollDirection.Down)
            caret.EnsureVisible()
            sv.ScrollViewportVerticallyByLine(Nautilus.Text.Editor.ScrollDirection.Up)
        End Sub

        Public Function GetCurrentSubName() As String
            Dim textView = EditorControl.TextView
            Dim text = textView.TextSnapshot
            Dim insertionIndex = textView.Caret.Position.TextInsertionIndex
            Dim lineNumber = text.GetLineFromPosition(insertionIndex).LineNumber

            For i = lineNumber To 0 Step -1
                Dim line = text.GetLineFromLineNumber(i)
                Dim Tokens = LineScanner.GetTokenEnumerator(line.GetText(), i)
                Dim token = Tokens.Current.Type
                If token = TokenType.Sub OrElse token = TokenType.Function Then
                    If Tokens.MoveNext() AndAlso Tokens.Current.Type = TokenType.Identifier Then
                        Return Tokens.Current.Text
                    End If

                ElseIf (token = TokenType.EndSub OrElse token = TokenType.EndFunction) AndAlso lineNumber <> i Then
                    Return ""
                End If
            Next

            Return ""
        End Function

        Sub FormatSub()
            _formatting = True
            Dim textView = _editorControl.TextView
            Dim insertionIndex = textView.Caret.Position.TextInsertionIndex
            Dim lineNumber = If(stopFormatingLine = -1, textView.TextSnapshot.GetLineFromPosition(insertionIndex).LineNumber, stopFormatingLine)
            CompilerService.FormatDocument(textView.TextBuffer, lineNumber, stopFormatingLine = -1)
            _formatting = False
            stopFormatingLine = -1
        End Sub

        Sub FormatLine(lineNum As Integer)
            _formatting = True
            Dim textView = _editorControl.TextView
            Dim lineText = textView.TextSnapshot.GetLineFromLineNumber(lineNum).GetText()

            Select Case LineScanner.GetFirstToken(lineText, lineNum).Type
                Case TokenType.Sub, TokenType.EndSub, TokenType.Function, TokenType.EndFunction
                    lineNum = -1 ' format the document
            End Select

            CompilerService.FormatDocument(textView.TextBuffer, lineNum, stopFormatingLine = -1)
            _formatting = False
            stopFormatingLine = -1
        End Sub

        Public Function GetEndSubPos(pos As Integer) As Integer
            Dim textView = EditorControl.TextView
            Dim text = textView.TextSnapshot
            Dim lineNumber = textView.TextSnapshot.GetLineFromPosition(pos).LineNumber + 1

            Dim line As ITextSnapshotLine
            For i = lineNumber To textView.TextSnapshot.LineCount - 1
                line = text.GetLineFromLineNumber(i)
                Dim token = LineScanner.GetFirstToken(line.GetText(), i)
                Select Case token.Type
                    Case TokenType.Sub, TokenType.Function
                        Return -1
                    Case TokenType.EndSub, TokenType.EndFunction
                        Return line.Start + token.Column
                End Select
            Next

            Return -1
        End Function

        Public Function FindEventHandler(name As String) As Integer
            Dim snapshot = EditorControl.TextView.TextSnapshot
            name = name.ToLower()
            For Each line In snapshot.Lines
                Dim code = line.GetText().Trim(" "c, vbTab)
                If code = "" Then Continue For

                Dim Tokens = LineScanner.GetFirstTokens(line.GetText(), line.LineNumber, 2)
                If Tokens.Count < 2 Then Continue For

                If Tokens(0).Type = TokenType.Sub OrElse Tokens(0).Type = TokenType.Function Then
                    If Tokens(1).Type = TokenType.Identifier Then
                        If Tokens(1).LCaseText = name Then Return line.Start + Tokens(1).Column
                    End If
                End If

            Next
            Return -1
        End Function

        Public Sub UpdateGlobalSubsList()
            _GlobalSubs.Clear()
            _GlobalSubs.Add(AddNewFunc)
            _GlobalSubs.Add(AddNewSub)

            Dim textView = EditorControl.TextView
            Dim text = textView.TextSnapshot
            Dim hasChanged = False

            For i = 0 To text.LineCount - 1
                Dim line = text.GetLineFromLineNumber(i)
                Dim lineText = line.GetText()
                Dim token = LineScanner.GetFirstToken(lineText, i)
                Dim type = token.Type
                If type = TokenType.Sub OrElse type = TokenType.Function Then
                    token = LineScanner.GetFirstToken(lineText.Substring(token.EndColumn), i)
                    If token.Type = TokenType.Identifier Then
                        Dim subName = token.Text
                        If Not _EventHandlers.ContainsKey(subName) Then
                            ' If name has the form Control_Event, add ot to EventHandlers.
                            Dim info = GetHandlerInfo(subName.ToLower())
                            If info.ControlName <> "" Then
                                _EventHandlers(subName) = info
                            Else
                                _GlobalSubs.Add(subName)
                                If Not hasChanged Then
                                    hasChanged = Not _ControlEvents.Contains(subName)
                                End If
                            End If
                        End If
                    End If
                End If
            Next

            If Not hasChanged AndAlso _ControlEvents.Count = _GlobalSubs.Count Then Return

            _GlobalSubs.Sort()
            If _MdiView Is Nothing OrElse CStr(_MdiView.CmbControlNames.SelectedItem) = "(Global)" Then
                If _MdiView IsNot Nothing Then _MdiView.FreezeCmbEvents = True
                _ControlEvents.Clear()
                For Each sb In _GlobalSubs
                    _ControlEvents.Add(sb)
                Next
                If _MdiView IsNot Nothing Then _MdiView.FreezeCmbEvents = False
            End If
        End Sub

        Public Function GetHandlerInfo(subName As String) As EventInformation
            If _ControlsInfo Is Nothing Then Return New EventInformation("", "")
            subName = subName.ToLower()
            For Each controlInfo In _ControlsInfo
                Dim controlName = controlInfo.Key
                If subName.StartsWith(controlName & "_") Then
                    Dim eventName = subName.Substring(controlName.Length + 1)
                    Dim events = PreCompiler.GetEvents(controlInfo.Value)
                    For Each ev In events
                        If ev.ToLower() = eventName Then
                            For Each name In _ControlNames
                                If name.ToLower = controlName Then Return New EventInformation(name, ev)
                            Next
                            Return New EventInformation("", "")
                        End If
                    Next
                End If
            Next

            Return New EventInformation("", "")
        End Function

        Public Sub RemoveBrokenHandlers()
            If _EventHandlers.Count = 0 Then Return

            Dim textView = EditorControl.TextView
            Dim text = textView.TextSnapshot
            Dim found As New List(Of String)

            For i = 0 To text.LineCount - 1
                Dim line = text.GetLineFromLineNumber(i)
                Dim Tokens = LineScanner.GetTokenEnumerator(line.GetText(), i)
                If Tokens.Current.Type = TokenType.Sub OrElse Tokens.Current.Type = TokenType.Function Then
                    If Tokens.MoveNext() AndAlso Tokens.Current.Type = TokenType.Identifier Then
                        Dim subName = Tokens.Current.Text
                        If _EventHandlers.ContainsKey(subName) Then
                            found.Add(subName)
                        End If
                    End If
                End If
            Next

            If _EventHandlers.Count > found.Count Then
                Dim handlers = _EventHandlers.Keys
                For i = handlers.Count - 1 To 0 Step -1
                    Dim h = handlers(i)
                    If Not found.Contains(h) Then
                        _EventHandlers.Remove(h)
                        If _EventHandlers.Count = found.Count Then Return
                    End If
                Next
            End If
        End Sub

        Friend Sub FixEventHandlers(oldName As String, newName As String)
            Dim keys = EventHandlers.Keys

            For i = keys.Count - 1 To 0 Step -1
                Dim oldHandler = keys(i)
                Dim eventInfo = EventHandlers(oldHandler)
                If eventInfo.ControlName.ToLower() = oldName.ToLower() Then
                    Dim newHandler = $"{newName}_{eventInfo.EventName}"
                    Dim pos = FindEventHandler(oldHandler)
                    _editorControl.EditorOperations.ReplaceText(New Span(pos, oldHandler.Length), newHandler, _undoHistory)
                    EventHandlers.Remove(oldHandler)
                    EventHandlers(newHandler) = New EventInformation(newName, eventInfo.EventName)
                End If
            Next

        End Sub

        Public Sub ShowErrors(errors As List(Of [Error]))
            ' don't use _errorListControl because it can be Nothing
            ErrorListControl.ErrorTokens.Clear()
            Me.Errors.Clear()

            For Each err As [Error] In errors
                Dim token = err.Token
                token.Line = err.Line - sVB.LineOffset
                ErrorListControl.ErrorTokens.Add(token)

                If err.Line = -1 Then
                    Me.Errors.Add(err.Description)
                Else
                    Me.Errors.Add($"{token.Line + 1},{err.Column + 1}: {err.Description}")
                End If
            Next

            _errorListControl.SelectError(0)

        End Sub

        Public Function CompileGlobalModule() As List(Of Parser)
            If Me.File = "" OrElse IsTheGlobalFile Then Return Nothing
            Dim inputDir = IO.Path.GetDirectoryName(Me.File)
            Dim outputFileName = sVB.GetOutputFileName(Me.File, False)
            Return sVB.CompileGlobalModule(inputDir, outputFileName)
        End Function

        Public Function GetFormNames(Optional normalize As Boolean = False) As List(Of String)
            Dim forms As New List(Of String)
            If Not Me.IsTheGlobalFile AndAlso Me.Form = "" Then Return forms

            Dim inputDir = Me.File
            If inputDir = "" Then Return forms

            inputDir = Path.GetDirectoryName(inputDir)

            For Each xamlFile In Directory.GetFiles(inputDir, "*.xaml")
                Dim name = DiagramHelper.Helper.GetFormNameFromXaml(xamlFile)
                If name = "" Then Continue For
                forms.Add(If(normalize, name.ToLower(), name))
            Next

            Return forms
        End Function
    End Class
End Namespace
