Imports System.ComponentModel
Imports System.ComponentModel.Composition
Imports System.IO
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Core.Undo.DataExports
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Operations

Namespace Microsoft.Windows.Documents
    Public Class FileDocument
        Implements INotifyPropertyChanged, INotifyImport

        Private _isDirty As Boolean
        Private saveMarker As UndoTransactionMarker
        Private _textBuffer As ITextBuffer
        Private textBufferUndoManager As ITextBufferUndoManager

        Public Property ContentType As String
            Get
                Return TextBuffer.ContentType
            End Get

            Set(value As String)
                TextBuffer.ContentType = value
                NotifyProperty("ContentType")
            End Set
        End Property

        Public ReadOnly Property FilePath As String

        Public Property IsDirty As Boolean
            Get
                Return _isDirty
            End Get

            Private Set(value As Boolean)
                If _isDirty <> value Then
                    _isDirty = value
                    NotifyProperty("IsDirty")
                    NotifyProperty("Title")
                End If
            End Set

        End Property

        Public ReadOnly Property IsNew As Boolean
            Get
                Return FilePath Is Nothing
            End Get
        End Property

        Public ReadOnly Property TextBuffer As ITextBuffer
            Get
                If _textBuffer Is Nothing Then CreateBuffer()

                Return _textBuffer
            End Get
        End Property

        <Import>
        Public Property TextBufferFactory As ITextBufferFactory

        <Import>
        Public Property TextBufferUndoManagerProvider As ITextBufferUndoManagerProvider

        Public ReadOnly Property Title As String
            Get
                Dim text As String = (If(IsDirty, "*", ""))
                If IsNew Then Return "Untitled" & text

                Return $"{Path.GetFileName(FilePath)}{text}"
            End Get
        End Property

        Public ReadOnly Property UndoHistory As UndoHistory

        <Import>
        Public Property UndoHistoryRegistry As IUndoHistoryRegistry

        Public Event PropertyChanged As PropertyChangedEventHandler Implements ComponentModel.INotifyPropertyChanged.PropertyChanged

        Public Sub New(filePath As String)
            If filePath IsNot Nothing AndAlso Not File.Exists(filePath) Then
                Throw New ArgumentException("The specified file path doesn't exist.")
            End If

            _FilePath = filePath
            If _FilePath IsNot Nothing AndAlso Not Path.IsPathRooted(_FilePath) Then
                _FilePath = Path.GetFullPath(filePath)
            End If

            saveMarker = New UndoTransactionMarker
        End Sub

        Public Sub Close()
            RemoveHandler TextBuffer.Changed, AddressOf TextBufferChanged
            RemoveHandler _UndoHistory.UndoRedoHappened, AddressOf UndoRedoHappened
        End Sub

        Private Sub CreateBuffer()
            If IsNew Then
                _textBuffer = New BufferFactory().CreateTextBuffer()
            Else
                Using reader As New StreamReader(FilePath)
                    _textBuffer = New BufferFactory().CreateTextBuffer(reader, GetContentTypeFromFileExtension())
                End Using
            End If

            AddHandler TextBuffer.Changed, AddressOf TextBufferChanged
        End Sub

        Private Function GetContentTypeFromFileExtension() As String
            Dim result As String = "text"

            Select Case Path.GetExtension(FilePath).ToLowerInvariant()

                Case ".m"
                    result = "text.m"

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

        Private Sub NotifyProperty(propertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
        End Sub

        Private Sub ImportCompleted() Implements ComponentModel.Composition.INotifyImport.ImportCompleted
            If Not UndoHistoryRegistry.TryGetHistory(TextBuffer, _UndoHistory) Then
                _UndoHistory = UndoHistoryRegistry.RegisterHistory(TextBuffer)
            End If

            _UndoHistory.ReplaceMarkerOnTop(saveMarker, True)
            AddHandler _UndoHistory.UndoRedoHappened, AddressOf UndoRedoHappened
            textBufferUndoManager = TextBufferUndoManagerProvider.GetTextBufferUndoManager(TextBuffer)
        End Sub

        Public Sub Save()
            If IsNew Then
                Throw New InvalidOperationException("Can't save a transient document.")
            End If

            Using writer As New StreamWriter(FilePath)
                TextBuffer.CurrentSnapshot.Write(writer)
            End Using

            _UndoHistory.ReplaceMarkerOnTop(saveMarker, True)
            IsDirty = False
        End Sub

        Public Sub SaveAs(filePath As String)
            Using writer As New StreamWriter(filePath)
                TextBuffer.CurrentSnapshot.Write(writer)
            End Using

            _FilePath = filePath
            NotifyProperty("FilePath")
            _UndoHistory.ReplaceMarkerOnTop(saveMarker, True)
            IsDirty = False
            NotifyProperty("Title")
        End Sub

        Private Sub TextBufferChanged(sender As Object, e As TextChangedEventArgs)
            IsDirty = True
        End Sub

        Private Sub UndoRedoHappened(sender As Object, e As UndoRedoEventArgs)
            Dim __ As Object = Nothing
            If _UndoHistory.TryFindMarkerOnTop(saveMarker, __) Then
                IsDirty = False
            Else
                IsDirty = True
            End If
        End Sub
    End Class
End Namespace
