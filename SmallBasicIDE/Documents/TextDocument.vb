Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Core.Undo.DataExports
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations
Imports Microsoft.SmallBasic.Shell
Imports Microsoft.Windows.Controls
Imports Microsoft.SmallBasic.WinForms
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
Imports Microsoft.SmallBasic.LanguageService
Imports System.Windows

Namespace Microsoft.SmallBasic.Documents
    Public Class TextDocument
        Inherits FileDocument
        Implements INotifyImport

        Private saveMarker As UndoTransactionMarker
        Private _textBuffer As ITextBuffer
        Private textBufferUndoManager As ITextBufferUndoManager
        Private _undoHistory As UndoHistory
        Private _editorControl As CodeEditorControl
        Private _errorListControl As ErrorListControl
        Private _errors As New ObservableCollection(Of String)()
        Private _caretPositionText As String
        Private _programDetails As Object
        Public Property BaseId As String

        Public Property WordHighlightColor As System.Windows.Media.Color = System.Windows.Media.Colors.LightGray

        Public Property MatchingPairsHighlightColor As System.Windows.Media.Color = System.Windows.Media.Colors.LightGreen


        Public Property CaretPositionText As String
            Get
                Return _caretPositionText
            End Get

            Private Set(ByVal value As String)
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

        Public ReadOnly Property Errors As ObservableCollection(Of String)
            Get
                Return _errors
            End Get
        End Property

        Public Property ProgramDetails As Object
            Get
                Return _programDetails
            End Get

            Set(ByVal value As Object)
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
                Dim currentSnapshot = TextBuffer.CurrentSnapshot
                Return currentSnapshot.GetText(0, currentSnapshot.Length)
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

                Return $"{Path.GetFileName(FilePath)}{text}"
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
        End Sub


        Private Sub OnCaretPositionChanged(sender As Object, e As CaretPositionChangedEventArgs)
            If IgnoreCaretPosChange Or StillWorking Then Return
            EditorControl.Dispatcher.BeginInvoke(
                    Sub()
                        StillWorking = True
                        Try
                            Dim handlerName = FindCurrentSub()
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
                                    reCompileSourceCode = True
                                    _editorControl.ClearHighlighting()
                                End If

                                HighlightMatchingPair()
                                sourceCodeChanged = False
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
                Dim token = LineScanner.GetToken(span.GetText().ToLower())
                Dim tokenType = LineScanner.GetTokenType(token)
                If tokenType = TokenType.Keyword Then
                    Dim pos = span.Start
                    Dim line = snapshot.GetLineFromPosition(pos)
                    HighlightBlockKeywords(line.LineNumber, snapshot, token)
                End If
                Return
            End If

            Dim index = If(sel.IsEmpty,
                    textView.Caret.Position.TextInsertionIndex,
                    sel.ActiveSnapshotSpan.Start
            )

            Dim matchingPair1 As Char
            Dim matchingPair2 As Char = ChrW(0)
            Dim direction As Integer = 1

            For pos = If(index = length, 1, 0) To If(index = 0 OrElse Not sel.IsEmpty, 0, 1)
                index -= pos
                Select Case snapshot(index)
                    Case "("c
                        matchingPair1 = "("c
                        matchingPair2 = ")"c
                    Case "["c
                        matchingPair1 = "["c
                        matchingPair2 = "]"c
                    Case "{"c
                        matchingPair1 = "{"c
                        matchingPair2 = "}"c
                    Case ")"c
                        matchingPair1 = ")"c
                        matchingPair2 = "("c
                        direction = -1
                    Case "]"c
                        matchingPair1 = "]"c
                        matchingPair2 = "["c
                        direction = -1
                    Case "}"c
                        matchingPair1 = "}"c
                        matchingPair2 = "{"c
                        direction = -1
                End Select
                If AscW(matchingPair2) <> 0 Then Exit For
            Next

            If AscW(matchingPair2) = 0 Then
                If sel.IsEmpty Then HighlightBlockKeywords()
                Return
            End If

            Dim i = index
            Dim pair1Count = 0
            Do
                i += direction
                If i = -1 OrElse i >= length Then Exit Do
                Select Case snapshot(i)
                    Case matchingPair1
                        pair1Count += 1
                    Case matchingPair2
                        If pair1Count = 0 Then
                            If index < i Then
                                _editorControl.HighlightWords(_MatchingPairsHighlightColor, (index, 1), (i, 1))
                            Else
                                _editorControl.HighlightWords(_MatchingPairsHighlightColor, (i, 1), (index, 1))
                            End If
                            Return
                        End If
                        pair1Count -= 1
                End Select
            Loop

        End Sub
        Public Sub HighlightEnclosingBlockKeywords()
            Dim textView = _editorControl.TextView
            Dim pos = textView.Caret.Position.TextInsertionIndex
            Dim lineNumber = textView.TextSnapshot.GetLineNumberFromPosition(pos)
            Dim statement = GetBlock(lineNumber, Nothing)
            _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))
        End Sub

        Private Sub HighlightBlockKeywords()
            Dim textView = _editorControl.TextView
            Dim snapshot = textView.TextSnapshot
            Dim index = textView.Caret.Position.TextInsertionIndex
            Dim line = snapshot.GetLineFromPosition(index)
            Dim pos = index - line.Start

            For Each tokenInfo In LineScanner.GetTokens(line.GetText(), line.LineNumber)
                If pos >= tokenInfo.Column AndAlso pos <= tokenInfo.EndColumn Then
                    If tokenInfo.TokenType = TokenType.Keyword Then
                        HighlightBlockKeywords(line.LineNumber, snapshot, tokenInfo.Token)
                    End If
                    Return
                End If
            Next
        End Sub

        Dim highlightCompiler As New Compiler()

        Private Sub HighlightBlockKeywords(lineNumber As Integer, snapshot As ITextSnapshot, token As Token)
            Select Case token
                Case Token.Sub, Token.EndSub, Token.Function, Token.EndFunction, Token.Return
                    Dim statement = GetBlock(lineNumber, GetType(Statements.SubroutineStatement))
                    _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))

                Case Token.If, Token.Then, Token.ElseIf, Token.Else, Token.EndIf
                    Dim statement = GetBlock(lineNumber, GetType(Statements.IfStatement))
                    _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))

                Case Token.For, Token.To, Token.Step, Token.Next, Token.EndFor
                    Dim statement = GetBlock(lineNumber, GetType(Statements.ForStatement))
                    _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))

                Case Token.While, Token.Wend, Token.EndWhile
                    Dim statement = GetBlock(lineNumber, GetType(Statements.WhileStatement))
                    _editorControl.HighlightWords(_WordHighlightColor, GetSpans(statement?.GetKeywords()))

                Case Token.ContinueLoop, Token.ExitLoop

            End Select
        End Sub

        Public Function GetBlock(lineNumber As Integer, statementType As Type) As Statements.Statement

            If reCompileSourceCode Then
                reCompileSourceCode = False
                highlightCompiler.Compile(New StringReader(Me.Text))
            End If

            Dim statement = Completion.CompletionHelper.GetStatement(highlightCompiler, lineNumber)

            If statement Is Nothing Then Return Nothing
            If statementType IsNot Nothing AndAlso statement.GetType() Is statementType Then Return statement

            statement = statement.GetStatement(lineNumber)

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

        Public Function GetSpans(tokens As List(Of TokenInfo)) As (Integer, Integer)()
            If tokens Is Nothing Then Return Nothing

            Dim spans As New List(Of (Start As Integer, Length As Integer))
            Dim snapshot = _editorControl.TextView.TextSnapshot
            Dim lineNum = -1
            Dim linestart = 0

            For Each token In tokens
                If lineNum <> token.Line Then
                    lineNum = token.Line
                    linestart = snapshot.GetLineFromLineNumber(token.Line).Start
                End If

                spans.Add((
                    linestart + token.Column,
                    token.EndColumn - token.Column
                ))
            Next

            Return spans.ToArray()
        End Function


        Sub UpdateCombos(eventInfo As (ControlName As String, EventName As String))
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
            If IsNew Then
                Throw New InvalidOperationException("Can't save a transient document.")
            End If

            Using stream As Stream = Open()
                stream.SetLength(0L)

                Using writer As StreamWriter = New StreamWriter(stream)
                    TextBuffer.CurrentSnapshot.Write(writer)
                End Using
            End Using

            _undoHistory.ReplaceMarkerOnTop(saveMarker, True)
            IsDirty = False
        End Sub

        Public Sub SaveAs(ByVal filePath As String)
            Dim filePath2 = MyBase.FilePath
            MyBase.FilePath = filePath

            Try
                Save()
            Catch
                MyBase.FilePath = filePath2
                Throw
            End Try
        End Sub

        Private Sub CreateBuffer()
            If IsNew Then
                _textBuffer = New BufferFactory().CreateTextBuffer()
            Else

                Using reader As StreamReader = New StreamReader(FilePath)
                    _textBuffer = New BufferFactory().CreateTextBuffer(reader, GetContentTypeFromFileExtension())
                End Using
            End If

            AddHandler _textBuffer.Changed, AddressOf OnTextBufferChanged
        End Sub

        Private Function GetContentTypeFromFileExtension() As String
            Dim result = "text"

            Select Case Path.GetExtension(FilePath).ToLowerInvariant()
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
        Dim reCompileSourceCode As Boolean

        Private Sub OnTextBufferChanged(sender As Object, e As TextChangedEventArgs)
            sourceCodeChanged = True
            IsDirty = True
            If _formatting OrElse _IgnoreCaretPosChange OrElse StillWorking Then Return

            StillWorking = True
            Try
                EditorControl.Dispatcher.BeginInvoke(
                  Sub()
                      StillWorking = True
                      UpdateGlobalSubsList()

                      ' re-format if lines changed
                      For Each change In e.Changes
                          If change.Delta > 2 AndAlso change.NewText.Contains(vbLf) AndAlso change.NewText.Trim() <> "" Then
                              FormatSub()
                              Exit For
                          End If

                          If change.OldText.Contains(vbLf) Then
                              FormatSub()
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
                        AutoCompleteBlock(textView, line, code, keyword, $"For  = ? To ? #   ", "Next", paran.Length)

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

            If endBlock <> "" AndAlso endBlock <> "EndSub" AndAlso endBlock <> "EndFunction" Then
                For i = line.LineNumber + 1 To text.LineCount - 1
                    Dim nextLine = text.GetLineFromLineNumber(i)
                    Dim LineCode = nextLine.GetText()
                    If LineCode.Trim(" "c, vbTab, "(").ToLower() <> "" Then
                        Dim indent_Keyword = leadingSpace & keyword
                        If LineCode = leadingSpace & endBlock Then
                            L += nextLine.Length + 2
                            Try
                                nextLine = text.GetLineFromLineNumber(i + 1)
                                LineCode = nextLine.GetText()
                                If LineCode.Trim(" "c, vbTab, "(").ToLower() = "" Then
                                    L += 2
                                End If
                            Catch ex As Exception

                            End Try
                        ElseIf LineCode.StartsWith(indent_KeyWord) Then
                            Exit For
                        Else
                            For j = i + 1 To text.LineCount - 1
                                nextLine = text.GetLineFromLineNumber(j)
                                LineCode = nextLine.GetText()
                                If LineCode.StartsWith(indent_KeyWord) Then Exit For
                                If LineCode = leadingSpace & endBlock Then
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

            Dim inden = leadingSpace.Length
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
            textView.Caret.MoveTo(line.Start + inden + Len(keyword) + 1 + n)
            stopFormatingLine = line.LineNumber
            _formatting = False
        End Sub

        Friend Sub Focus(Optional moveToStart As Boolean = False)
            Dim txtView = CType(EditorControl.TextView, AvalonTextView)
            txtView.VisualElement.Focus()
            If moveToStart Then
                txtView.Caret.MoveTo(0)
            End If
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
        Public Property EventHandlers As New Dictionary(Of String, (ControlName As String, EventName As String))

        Dim _form As String
        Public Property Form As String
            Get
                Return _form
            End Get

            Set(value As String)
                _form = value
                NotifyProperty("Title")
            End Set
        End Property

        Public Property ControlsInfo As Dictionary(Of String, String)


        Function GenerateCodeBehind(formDesigner As DiagramHelper.Designer, updateControlInfo As Boolean) As String
            If _form = "" Then Return ""

            Dim hint As New Text.StringBuilder
            Dim declaration As New Text.StringBuilder
            Dim formName = _form
            hint.AppendLine("'@Form Hints:")
            hint.AppendLine($"'#{formName}{{")
            Dim controlsInfoList As New Dictionary(Of String, String)
            Dim controlNamesList As New List(Of String)
            controlNamesList.Add("(Global)")

            controlsInfoList(formName.ToLower()) = "Form"
            controlsInfoList("me") = "Form"
            controlNamesList.Add(formName)
            declaration.AppendLine($"Me = ""{formName}""")

            For Each c As UIElement In formDesigner.Items
                Dim name = formDesigner.GetControlNameOrDefault(c)

                If name <> "" Then
                    Dim typeName = PreCompiler.GetModuleName(c.GetType().Name)
                    controlsInfoList(name.ToLower()) = typeName
                    controlNamesList.Add(name)
                    hint.AppendLine($"'    {name}: {typeName}")
                    declaration.AppendLine($"{name} = ""{name.ToLower()}""")
                End If
            Next

            hint.AppendLine("'}")
            hint.AppendLine()
            hint.Append(declaration)
            'hint.AppendLine($"Forms.AppPath = ""{xamlPath}""")
            hint.AppendLine($"{formName} = Forms.LoadForm(""{formName}"", ""{Path.GetFileNameWithoutExtension(Me.FilePath)}.xaml"")")
            If formDesigner.AllowTransparency Then
                hint.AppendLine($"Form.AllowTransparency({formName})")
            End If
            hint.AppendLine($"Form.Show({formName})")

            ' Remove Hamdlers of deleted or renamed controls
            For i = _EventHandlers.Count - 1 To 0 Step -1
                If Not controlNamesList.Contains(_EventHandlers.Values(i).ControlName) Then
                    _EventHandlers.Remove(_EventHandlers.Keys(i))
                End If
            Next

            ' Remove handlers of renamed subs
            RemoveBrokenHandlers()

            If EventHandlers.Count > 0 Then
                hint.AppendLine()
                hint.AppendLine("'#Events{")
                Dim sbHandlers As New Text.StringBuilder

                Dim controlEvents = From eventHandler In EventHandlers
                                    Group By eventHandler.Value.ControlName
                           Into EventInfo = Group

                For Each ev In controlEvents
                    hint.Append($"'    {ev.ControlName}:")
                    sbHandlers.AppendLine($"' {ev.ControlName} Events:")
                    sbHandlers.AppendLine($"Control.HandleEvents({formName}, {ev.ControlName})")
                    For Each info In ev.EventInfo
                        Dim controlName = ev.ControlName.ToLower
                        If controlsInfoList.ContainsKey(controlName) Then
                            hint.Append($" {info.Value.EventName}")
                            Dim module1 = controlsInfoList(controlName)
                            Dim moduleName = PreCompiler.GetEventModule(module1, info.Value.EventName)
                            sbHandlers.AppendLine($"{moduleName}.{info.Value.EventName} = {info.Key}")
                        End If
                    Next
                    hint.AppendLine()
                    sbHandlers.AppendLine()
                Next
                hint.AppendLine("'}")
                hint.AppendLine()
                hint.AppendLine(sbHandlers.ToString())
            End If

            If updateControlInfo Then
                _form = formName
                _ControlsInfo = controlsInfoList
                AddProperty("ControlsInfo", _ControlsInfo)
                AddProperty("ControlNames", _ControlNames)

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

            Return hint.ToString()
        End Function

        Sub AddProperty(name As String, value As Object)
            Try
                TextBuffer.Properties.AddProperty(name, value)
            Catch
                TextBuffer.Properties.RemoveProperty(name)
                TextBuffer.Properties.AddProperty(name, value)
            End Try
        End Sub

        Public Function GetCodeBehind(Optional ToCompile As Boolean = False) As String
            Dim codeFileHasHints = Me.Text.Contains("'@Form Hints:")

            Dim genCodefile = FilePath?.Substring(0, FilePath.Length - 2) + "sb.gen"
            If genCodefile = "" OrElse Not File.Exists(genCodefile) Then
                Return If(ToCompile, "", Me.Text)
            Else
                If ToCompile Then
                    Return File.ReadAllText(genCodefile)
                Else
                    Return If(codeFileHasHints, Me.Text, File.ReadAllText(genCodefile))
                End If
            End If

        End Function

        Public Sub ParseFormHints()
            Dim code = GetCodeBehind()

            Dim info = PreCompiler.ParseFormHints(code)
            If info Is Nothing Then Return

            Me.Form = info.Form
            Me.ControlsInfo = info.ControlsInfo
            Me.ControlsInfo.Add("me", "Form")

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
            AddProperty("ControlNames", _ControlNames)
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
        Private Const AddNewFunc As String  = "(Add New Function)"
        Private ReadOnly globalFunc As String = vbCrLf & "
'------------------------------------------------
Function #( )
   
   Return 0
EndFunction
"
        Public Property IgnoreCaretPosChange As Boolean

        Public Function AddEventHandler(controlName As String, eventName As String) As Boolean
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
                    _EventHandlers(handlerName) = (controlName, eventName)
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
                    SelectCurrentWord()
                    _ControlEvents.Add(handlerName)

                Else
                    _EventHandlers(handlerName) = (controlName, eventName)
                    Dim handler = eventHandlerSub.Replace("#", handlerName)
                    _editorControl.EditorOperations.InsertText(handler, _undoHistory)
                    EnsureAtTop(Text.Length - 10)
                End If

            Else
                alreadyExists = True
                caret.MoveTo(pos)
                SelectCurrentWord()
            End If

            Me.Focus()

            Me.MdiView.CmbControlNames.SelectedItem = controlName
            Me.MdiView.SelectEventName(If(isGlobal, handlerName, eventName))
            _IgnoreCaretPosChange = False

            Return Not alreadyExists
        End Function

        Public Sub SelectWordAt(line As Integer, column As Integer, Optional viewAtTop As Boolean = True)
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

        Public Function FindCurrentSub() As String
            Dim textView = EditorControl.TextView
            Dim text = textView.TextSnapshot
            Dim insertionIndex = textView.Caret.Position.TextInsertionIndex
            Dim lineNumber = text.GetLineFromPosition(insertionIndex).LineNumber

            For i = lineNumber To 0 Step -1
                Dim line = text.GetLineFromLineNumber(i)
                Dim Tokens = LineScanner.GetTokenEnumerator(line.GetText(), i)
                Dim token = Tokens.Current.Token
                If token = Token.Sub OrElse token = Token.Function Then
                    If Tokens.MoveNext() AndAlso Tokens.Current.Token = Token.Identifier Then
                        Return Tokens.Current.Text
                    End If

                ElseIf (token = Token.EndSub OrElse token = Token.EndFunction) AndAlso lineNumber <> i Then
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

        Public Function FindEndSub(pos As Integer) As Integer
            Dim textView = EditorControl.TextView
            Dim text = textView.TextSnapshot
            Dim lineNumber = textView.TextSnapshot.GetLineFromPosition(pos).LineNumber + 1

            Dim line As ITextSnapshotLine
            For i = lineNumber To textView.TextSnapshot.LineCount - 1
                line = text.GetLineFromLineNumber(i)
                Dim tokenInfo = LineScanner.GetFirstToken(line.GetText(), i)
                Select Case tokenInfo.Token
                    Case Token.Sub, Token.Function
                        Return -1
                    Case Token.EndSub, Token.EndFunction
                        Return line.Start + tokenInfo.Column
                End Select
            Next

            Return -1
        End Function

        Public Function FindEventHandler(name As String) As Integer
            Dim text = EditorControl.TextView.TextSnapshot
            name = name.ToLower()
            For Each line In text.Lines
                Dim code = line.GetText().Trim(" "c, vbTab)
                If code = "" Then Continue For

                Dim Tokens = LineScanner.GetTokenEnumerator(line.GetText(), line.LineNumber)
                If Tokens.Current.Token = Token.Sub OrElse Tokens.Current.Token = Token.Function Then
                    If Tokens.MoveNext() AndAlso Tokens.Current.Token = Token.Identifier Then
                        If Tokens.Current.NormalizedText = name Then Return line.Start + Tokens.Current.Column
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
                Dim tokenInfo = LineScanner.GetFirstToken(lineText, i)
                Dim token = tokenInfo.Token
                If token = token.Sub OrElse token = token.Function Then
                    tokenInfo = LineScanner.GetFirstToken(lineText.Substring(tokenInfo.EndColumn), i)
                    If tokenInfo.Token = token.Identifier Then
                        Dim subName = tokenInfo.Text
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

        Public Function GetHandlerInfo(subName As String) As (ControlName As String, EventName As String)
            If _ControlsInfo Is Nothing Then Return ("", "")
            subName = subName.ToLower()
            For Each controlInfo In _ControlsInfo
                Dim controlName = controlInfo.Key
                If subName.StartsWith(controlName & "_") Then
                    Dim eventName = subName.Substring(controlName.Length + 1)
                    Dim events = PreCompiler.GetEvents(controlInfo.Value)
                    For Each ev In events
                        If ev.ToLower() = eventName Then
                            For Each name In _ControlNames
                                If name.ToLower = controlName Then Return (name, ev)
                            Next
                            Return ("", "")
                        End If
                    Next
                End If
            Next

            Return ("", "")
        End Function

        Public Sub RemoveBrokenHandlers()
            If _EventHandlers.Count = 0 Then Return

            Dim textView = EditorControl.TextView
            Dim text = textView.TextSnapshot
            Dim found As New List(Of String)

            For i = 0 To text.LineCount - 1
                Dim line = text.GetLineFromLineNumber(i)
                Dim Tokens = LineScanner.GetTokenEnumerator(line.GetText(), i)
                If Tokens.Current.Token = Token.Sub OrElse Tokens.Current.Token = Token.Function Then
                    If Tokens.MoveNext() AndAlso Tokens.Current.Token = Token.Identifier Then
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


    End Class
End Namespace
