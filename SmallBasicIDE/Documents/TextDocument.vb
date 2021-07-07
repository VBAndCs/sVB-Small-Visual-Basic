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

        Public Property ContentType As String
            Get
                Return TextBuffer.ContentType
            End Get
            Set(ByVal value As String)
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
                                    _editorControl.TextView.VisualElement.Focus()
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

                If IsNew Then
                    If Equals(BaseId, Nothing) Then
                        Return ResourceHelper.GetString("Untitled") & text
                    End If

                    Return String.Format(CultureInfo.CurrentCulture, ResourceHelper.GetString("ImportedProgram"), New Object(0) {BaseId}) & text
                End If

                Return $"{Path.GetFileName(FilePath)}{text} - {Path.GetFullPath(FilePath)}"
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

        Public Sub New(ByVal filePath As String)
            MyBase.New(filePath)
            saveMarker = New UndoTransactionMarker()
            App.GlobalDomain.AddComponent(Me)
            App.GlobalDomain.Bind()
        End Sub

        Private Sub OnCaretPositionChanged(sender As Object, e As CaretPositionChangedEventArgs)
            UpdateCaretPositionText()

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
            Catch __unusedException1__ As Exception
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
            UpdateCaretPositionText()
        End Sub

        Private Sub AutoCompleteBlocks(sender As Object, e As System.Windows.Input.KeyEventArgs)
            Dim textView = _editorControl.TextView
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
                        AutoCompleteBlock(textView, line, code, keyword, $"ElseIf {paran} Then#   ", "EndIf", paran.Length)
                    Case "for"
                        AutoCompleteBlock(textView, line, code, keyword, $"For #   ", "EndFor", paran.Length)
                    Case "while"
                        AutoCompleteBlock(textView, line, code, keyword, $"While {paran}#   ", "EndWhile", paran.Length)
                    Case "sub"
                        AutoCompleteBlock(textView, line, code, keyword, $"Sub #   ", "EndSub", paran.Length)
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
            Dim L = line.Length
            Dim text = textView.TextSnapshot
            Dim addBlockEnd = True

            If endBlock <> "" Then
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

            Dim nl = $"{vbCrLf}{leadingSpace}"
            _editorControl.EditorOperations.ReplaceText(
                 New Span(line.Start + start, L - start),
                block.Replace("#", nl) & If(addBlockEnd, nl & endBlock & nl, ""), _undoHistory)
            textView.Caret.MoveTo(line.Start + start + Len(keyword) + 1 + n)
        End Sub

        Private Sub UndoRedoHappened(sender As Object, e As UndoRedoEventArgs)
            Dim __ As Object = Nothing

            If _undoHistory.TryFindMarkerOnTop(saveMarker, __) Then
                IsDirty = False
            Else
                IsDirty = True
            End If
        End Sub

        Private Sub UpdateCaretPositionText()
            If _editorControl IsNot Nothing Then
                _editorControl.Dispatcher.BeginInvoke(
                            Sub()
                                Dim textView = _editorControl.TextView
                                Dim textInsertionIndex = textView.Caret.Position.TextInsertionIndex
                                Dim line = textView.TextSnapshot.GetLineFromPosition(textInsertionIndex)
                                CaretPositionText = $"{line.LineNumber + 1},{textInsertionIndex - line.Start + 1}"
                            End Sub,
                            DispatcherPriority.ContextIdle)
            End If
        End Sub

        Public Sub ImportCompleted() Implements INotifyImport.ImportCompleted
            OnBindCompleted()
        End Sub

        Dim _form As String
        Public Property Form As String
            Get
                Return _form
            End Get

            Set(value As String)
                If _form = value Then Return
                _form = value
                TextBuffer.Properties.AddProperty("FormName", _form)
            End Set
        End Property

        Dim _ControlsInfo As Collections.Generic.Dictionary(Of String, String)
        Public Property ControlsInfo As Collections.Generic.Dictionary(Of String, String)
            Get
                Return _ControlsInfo
            End Get

            Set(value As Collections.Generic.Dictionary(Of String, String))
                _ControlsInfo = value
                Try
                    TextBuffer.Properties.AddProperty("ControlsInfo", _ControlsInfo)
                Catch
                    TextBuffer.Properties.RemoveProperty("ControlsInfo")
                    TextBuffer.Properties.AddProperty("ControlsInfo", _ControlsInfo)
                End Try
            End Set
        End Property

        Public Function ParseFormHints(code As String) As Boolean
            Dim info = PreCompiler.ParseFormHints(code)
            If info Is Nothing Then Return False

            Me.Form = info.Form
            ControlsInfo = info.ControlsInfo
            Return True
        End Function

        Friend Shared Function FromCode(code As String) As TextDocument
            Dim n = New Random().Next(1, 1000000)
            Dim filename = Path.Combine(Path.GetTempPath(), $"file{n}.sb")
            Global.My.Computer.FileSystem.WriteAllText(filename, code, False)
            Return New TextDocument(filename)
        End Function
    End Class
End Namespace
