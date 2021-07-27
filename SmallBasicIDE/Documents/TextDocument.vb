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

                If _textBuffer Is Nothing Then
                    CreateBuffer()
                End If

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

                If _Form <> "" Then Return $"{_Form}{text}"

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
            _GlobalSubs.Add("(Add New Sub)")
            _ControlEvents.Add("(Add New Sub)")
            ParseFormHints()
            UpdateGlobalSubsList()
        End Sub

        Private Sub OnCaretPositionChanged(sender As Object, e As CaretPositionChangedEventArgs)
            If IgnoreCaretPosChange Or StillWorking Then Return
            EditorControl.Dispatcher.BeginInvoke(
                    Sub()
                        StillWorking = True
                        Try
                            Dim handlerName = FindCurrentEventHandler()
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
                            UpdateCaretPositionText()

                        Finally
                            StillWorking = False
                        End Try
                    End Sub,
                 DispatcherPriority.ContextIdle)

        End Sub

        Sub UpdateCombos(eventInfo As (ControlName As String, EventName As String))
            If CStr(_MdiView.CmbControlNames.SelectedItem) <> eventInfo.ControlName Then
                _MdiView.CmbControlNames.SelectedItem = eventInfo.ControlName
            End If

            _MdiView.SelectEventName(eventInfo.EventName)
        End Sub

        Public Overrides Sub Close()
            RemoveHandler _editorControl.TextView.Caret.PositionChanged, AddressOf OnCaretPositionChanged
            RemoveHandler _textBuffer.Changed, AddressOf TextBufferChanged
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

            AddHandler _textBuffer.Changed, AddressOf TextBufferChanged
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

        Private Sub TextBufferChanged(sender As Object, e As TextChangedEventArgs)
            IsDirty = True
            If _IgnoreCaretPosChange Or StillWorking Then Return

            StillWorking = True

            Try
                EditorControl.Dispatcher.BeginInvoke(
                  Sub()
                      StillWorking = True
                      UpdateGlobalSubsList()
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
            Dim textInsertionIndex = textView.Caret.Position.TextInsertionIndex
            If textInsertionIndex = 0 Then Return

            Dim c = text.GetText(textInsertionIndex - 1, 1)
            If Char.IsLetterOrDigit(c) OrElse c = "_" Then Return
            Dim line = text.GetLineFromPosition(textInsertionIndex)
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

                Dim keyword = code.Trim(" "c, vbTab, "(").ToLower()
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

            Catch ex As Exception
            End Try
        End Sub

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
            Dim iden = Indentation.CalculateIndentation(textView.TextSnapshot, line.LineNumber)
            Dim L = line.Length
            Dim text = textView.TextSnapshot
            Dim addBlockEnd = True

            If endBlock <> "" AndAlso endBlock <> "EndSub" AndAlso endBlock <> "EndFunction" Then
                For i = line.LineNumber + 1 To text.LineCount - 1
                    Dim nextLine = text.GetLineFromLineNumber(i)
                    Dim LineCode = nextLine.GetText()
                    If LineCode.Trim(" "c, vbTab, "(").ToLower() <> "" Then
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
                        ElseIf LineCode.StartsWith(leadingSpace & keyword) Then
                            Exit For
                        Else
                            For j = i + 1 To text.LineCount - 1
                                nextLine = text.GetLineFromLineNumber(j)
                                LineCode = nextLine.GetText()
                                If LineCode.StartsWith(leadingSpace & keyword) Then Exit For
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

            Dim nl = $"{vbCrLf}{Space(iden)}"
            EditorControl.EditorOperations.ReplaceText(
                 New Span(line.Start, L),
                Space(iden) & block.Replace("#", nl) & If(addBlockEnd, nl & endBlock & nl, ""), _undoHistory)
            textView.Caret.MoveTo(line.Start + iden + Len(keyword) + 1 + n)
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
            If _Form = "" Then Return ""

            Dim hint As New Text.StringBuilder
            Dim declaration As New Text.StringBuilder
            Dim formName = _Form
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
            hint.AppendLine("True = ""True""")
            hint.AppendLine("False = ""False""")
            'hint.AppendLine($"Forms.AppPath = ""{xamlPath}""")
            hint.AppendLine($"{formName} = Forms.LoadForm(""{formName}"", ""{Path.GetFileNameWithoutExtension(Me.FilePath)}.xaml"")")
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
                _Form = formName
                _ControlsInfo = controlsInfoList
                AddProperty("ControlsInfo", _ControlsInfo)
                AddProperty("ControlNames", _ControlNames)

                ' Note that ControlNames Property is bound to a cono box, so keep the existing collection
                controlNamesList.Sort()
                _ControlNames.Clear()
                For Each controlName In controlNamesList
                    _ControlNames.Add(controlName)
                Next
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

        End Sub

        Friend Shared Function FromCode(code As String) As TextDocument
            Dim n = New Random().Next(1, 1000000)
            Dim filename = Path.Combine(Path.GetTempPath(), $"file{n}.sb")
            Global.My.Computer.FileSystem.WriteAllText(filename, code, False)
            Return New TextDocument(filename)
        End Function

        Private ReadOnly eventHandlerSub As String = vbCrLf & "
'----------------------------
Sub #
   
EndSub
"

        Public Property IgnoreCaretPosChange As Boolean

        Public Function AddEventHandler(controlName As String, eventName As String) As Boolean
            Dim alreadyExists = False
            Dim isGlobal = (controlName = "(Global)")
            Dim handlerName = If(isGlobal, "", controlName & "_") & If(eventName = "(Add New Sub)", "", eventName)
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
                    Dim NewName = "NewSub_"
                    Dim n = 0
                    Try
                        n = Aggregate s In _GlobalSubs
                                 Where s.StartsWith(NewName)
                                 Let x = s.Substring(7)
                                 Into Max(If(IsNumeric(x), CInt(x), 0))
                    Catch ex As Exception
                    End Try

                    handlerName = NewName & n + 1
                    Dim handler = eventHandlerSub.Replace("#", handlerName)
                    _GlobalSubs.Add(handlerName)
                    _editorControl.EditorOperations.InsertText(handler, _undoHistory)
                    caret.MoveTo(Text.Length - 15)
                    SelectCurrentWord()
                    _ControlEvents.Add(handlerName)
                Else
                    _EventHandlers(handlerName) = (controlName, eventName)
                    Dim handler = eventHandlerSub.Replace("#", handlerName)
                    _editorControl.EditorOperations.InsertText(handler, _undoHistory)
                    caret.MoveTo(Text.Length - 10)
                End If
            Else
                alreadyExists = True
                caret.MoveTo(pos)
                SelectCurrentWord()
            End If

            caret.EnsureVisible()
            Me.Focus()

            Me.MdiView.CmbControlNames.SelectedItem = controlName
            Me.MdiView.SelectEventName(If(isGlobal, handlerName, eventName))
            _IgnoreCaretPosChange = False

            Return Not alreadyExists
        End Function

        Public Sub SelectWordAt(line As Integer, column As Integer)
            If line < 0 Then Return

            Dim currentSnapshot = _textBuffer.CurrentSnapshot
            If line < currentSnapshot.LineCount Then
                Dim lineSnap = currentSnapshot.GetLineFromLineNumber(line)

                If column < lineSnap.LengthIncludingLineBreak Then
                    Dim charIndex = lineSnap.Start + column
                    _editorControl.EditorOperations.ResetSelection()
                    _editorControl.TextView.Caret.MoveTo(charIndex)
                    SelectCurrentWord()
                End If
            End If
        End Sub

        Public Sub SelectCurrentWord()
            Dim caret = _editorControl.TextView.Caret
            Dim ops = _editorControl.EditorOperations
            Dim sv = _editorControl.TextView.ViewScroller
            ops.ResetSelection()
            caret.EnsureVisible()
            sv.ScrollViewportVerticallyByPage(Nautilus.Text.Editor.ScrollDirection.Down)
            caret.EnsureVisible()
            sv.ScrollViewportVerticallyByLine(Nautilus.Text.Editor.ScrollDirection.Up)
            ops.SelectCurrentWord()
            Focus()
        End Sub

        Public Function FindCurrentEventHandler() As String
            Dim textView = EditorControl.TextView
            Dim text = textView.TextSnapshot
            Dim insertionIndex = textView.Caret.Position.TextInsertionIndex
            Dim lineNumber = textView.TextSnapshot.GetLineFromPosition(insertionIndex).LineNumber

            For i = lineNumber To 0 Step -1
                Dim line = text.GetLineFromLineNumber(i)
                Dim Tokens = New LineScanner().GetTokenList(line.GetText(), i)
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

        Public Function FindEndSub(pos As Integer) As Integer
            Dim textView = EditorControl.TextView
            Dim text = textView.TextSnapshot
            Dim lineNumber = textView.TextSnapshot.GetLineFromPosition(pos).LineNumber + 1

            Dim line As ITextSnapshotLine
            For i = lineNumber To textView.TextSnapshot.LineCount - 1
                line = text.GetLineFromLineNumber(i)
                Dim Tokens = New LineScanner().GetTokenList(line.GetText(), i)
                Select Case Tokens.Current.Token
                    Case Token.Sub, Token.Function
                        Return -1
                    Case Token.EndSub, Token.EndFunction
                        Return line.Start + Tokens.Current.Column
                End Select
            Next

            Return -1
        End Function

        Public Function FindEventHandler(name As String) As Integer
            Dim text = EditorControl.TextView.TextSnapshot

            For Each line In text.Lines
                Dim code = line.GetText().Trim(" "c, vbTab)
                If code = "" Then Continue For

                Dim Tokens = New LineScanner().GetTokenList(line.GetText(), line.LineNumber)
                If Tokens.Current.Token = Token.Sub OrElse Tokens.Current.Token = Token.Function Then
                    If Tokens.MoveNext() AndAlso Tokens.Current.Token = Token.Identifier Then
                        If Tokens.Current.Text = name Then Return line.Start + Tokens.Current.Column
                    End If
                End If

            Next
            Return -1
        End Function


        Public Sub UpdateGlobalSubsList()
            _GlobalSubs.Clear()
            _GlobalSubs.Add("(Add New Sub)")

            Dim textView = EditorControl.TextView
            Dim text = textView.TextSnapshot

            For i = 0 To text.LineCount - 1
                Dim line = text.GetLineFromLineNumber(i)
                Dim Tokens = New LineScanner().GetTokenList(line.GetText(), i)
                If Tokens.Current.Token = Token.Sub OrElse Tokens.Current.Token = Token.Function Then
                    If Tokens.MoveNext() AndAlso Tokens.Current.Token = Token.Identifier Then
                        Dim subName = Tokens.Current.Text
                        If Not _EventHandlers.ContainsKey(subName) Then
                            ' If name has the form Control_Event, add ot to EventHandlers.
                            Dim info = GetHandlerInfo(subName.ToLower())
                            If info.ControlName <> "" Then
                                _EventHandlers(subName) = info
                            Else
                                _GlobalSubs.Add(subName)
                            End If
                        End If
                    End If
                End If
            Next

            _GlobalSubs.Sort()
            If _MdiView Is Nothing OrElse CStr(_MdiView.CmbControlNames.SelectedItem) = "(Global)" Then
                _ControlEvents.Clear()
                For Each sb In _GlobalSubs
                    _ControlEvents.Add(sb)
                Next
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
                Dim Tokens = New LineScanner().GetTokenList(line.GetText(), i)
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
                For Each ev In _EventHandlers.Keys
                    If Not found.Contains(ev) Then
                        _EventHandlers.Remove(ev)
                        If _EventHandlers.Count = found.Count Then Return
                    End If
                Next
            End If
        End Sub


    End Class
End Namespace
