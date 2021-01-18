Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Core.Undo.DataExports
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations
Imports Microsoft.SmallBasic.Shell
Imports Microsoft.Windows.Controls
Imports SmallBasic.WinForms
Imports System
Imports System.Collections.ObjectModel
Imports System.ComponentModel.Composition
Imports System.Globalization
Imports System.IO
Imports System.Windows.Media
Imports System.Windows.Threading

Namespace Microsoft.SmallBasic.Documents
    Public Class TextDocument
        Inherits FileDocument
        Implements INotifyImport

        Private saveMarker As UndoTransactionMarker
        Private textBufferField As ITextBuffer
        Private textBufferUndoManager As ITextBufferUndoManager
        Private undoHistoryField As UndoHistory
        Private editorControlField As CodeEditorControl
        Private errorListControlField As ErrorListControl
        Private errorsField As ObservableCollection(Of String) = New ObservableCollection(Of String)()
        Private caretPositionTextField As String
        Private _programDetails As Object
        Public Property BaseId As String

        Public Property CaretPositionText As String
            Get
                Return caretPositionTextField
            End Get
            Private Set(ByVal value As String)
                caretPositionTextField = value
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
                Return errorsField
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

                If textBufferField Is Nothing Then
                    CreateBuffer()
                End If

                Return textBufferField
            End Get
        End Property

        Public ReadOnly Property EditorControl As CodeEditorControl
            Get

                If editorControlField Is Nothing Then
                    editorControlField = New CodeEditorControl With {
                        .TextBuffer = TextBuffer
                    }
                    App.GlobalDomain.AddComponent(editorControlField)
                    App.GlobalDomain.Bind()
                    editorControlField.HighlightSearchHits = True
                    editorControlField.ScaleFactor = 1.0
                    editorControlField.Background = Brushes.Transparent
                    editorControlField.EditorOperations.TabSize = 2
                    editorControlField.IsLineNumberMarginVisible = True
                    editorControlField.Focus()
                    AddHandler editorControlField.TextView.Caret.PositionChanged, AddressOf OnCaretPositionChanged
                    editorControlField.Dispatcher.BeginInvoke(DispatcherPriority.Render, CType(Function()
                                                                                                   editorControlField.TextView.VisualElement.Focus()
                                                                                                   UpdateCaretPositionText()
                                                                                                   Return Nothing
                                                                                               End Function, DispatcherOperationCallback), Nothing)
                End If

                Return editorControlField
            End Get
        End Property

        Public ReadOnly Property ErrorListControl As ErrorListControl
            Get

                If errorListControlField Is Nothing Then
                    errorListControlField = New ErrorListControl(Me)
                    errorListControlField.ItemsSource = Errors
                End If

                Return errorListControlField
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
                Return undoHistoryField
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

        Private Sub OnCaretPositionChanged(ByVal sender As Object, ByVal e As CaretPositionChangedEventArgs)
            UpdateCaretPositionText()
        End Sub

        Public Overrides Sub Close()
            RemoveHandler editorControlField.TextView.Caret.PositionChanged, AddressOf OnCaretPositionChanged
            RemoveHandler textBufferField.Changed, AddressOf TextBufferChanged
            RemoveHandler undoHistoryField.UndoRedoHappened, AddressOf UndoRedoHappened
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

            undoHistoryField.ReplaceMarkerOnTop(saveMarker, True)
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
                textBufferField = New BufferFactory().CreateTextBuffer()
            Else

                Using reader As StreamReader = New StreamReader(FilePath)
                    textBufferField = New BufferFactory().CreateTextBuffer(reader, GetContentTypeFromFileExtension())
                End Using
            End If

            AddHandler textBufferField.Changed, AddressOf TextBufferChanged
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
            If Not UndoHistoryRegistry.TryGetHistory(TextBuffer, undoHistoryField) Then
                undoHistoryField = UndoHistoryRegistry.RegisterHistory(TextBuffer)
            End If

            undoHistoryField.ReplaceMarkerOnTop(saveMarker, True)
            AddHandler undoHistoryField.UndoRedoHappened, AddressOf UndoRedoHappened
            textBufferUndoManager = TextBufferUndoManagerProvider.GetTextBufferUndoManager(TextBuffer)
        End Sub

        Private Sub TextBufferChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs)
            IsDirty = True
            UpdateCaretPositionText()
        End Sub

        Private Sub UndoRedoHappened(ByVal sender As Object, ByVal e As UndoRedoEventArgs)
            Dim __ As Object = Nothing

            If undoHistoryField.TryFindMarkerOnTop(saveMarker, __) Then
                IsDirty = False
            Else
                IsDirty = True
            End If
        End Sub

        Private Sub UpdateCaretPositionText()
            If editorControlField IsNot Nothing Then
                editorControlField.Dispatcher.BeginInvoke(
                            Sub()
                                If editorControlField IsNot Nothing Then
                                    Dim textInsertionIndex = editorControlField.TextView.Caret.Position.TextInsertionIndex
                                    Dim lineFromPosition = editorControlField.TextView.TextSnapshot.GetLineFromPosition(textInsertionIndex)
                                    CaretPositionText = $"{lineFromPosition.LineNumber + 1},{textInsertionIndex - lineFromPosition.Start + 1}"
                                End If
                            End Sub,
                            DispatcherPriority.ContextIdle)
            End If
        End Sub

        Public Sub ImportCompleted() Implements INotifyImport.ImportCompleted
            OnBindCompleted()
        End Sub


        Public ReadOnly Property Form As String
        Public ReadOnly Property ControlsInfo As New Collections.Generic.Dictionary(Of String, String)

        Public Function ParseFormHints(code As String) As Boolean
            Dim info = PreCompiler.ParseFormHints(code)
            If info Is Nothing Then Return False

            Me._Form = info.Form
            _ControlsInfo = info.ControlsInfo
            Return True
        End Function

        Friend Shared Function FromCode(code As String) As TextDocument
            Dim n = New Random().Next(1, 1000000)
            Dim filename = Path.Combine(Path.GetTempPath(), $"file{n}.sb")
            My.Computer.FileSystem.WriteAllText(filename, code, False)
            Return New TextDocument(filename)
        End Function
    End Class
End Namespace
