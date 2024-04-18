Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text
    Friend MustInherit Class BaseBuffer
        Implements ITextBuffer, IPropertyOwner

        Private NotInheritable Class Edit
            Implements ITextEdit, IDisposable

            Private _baseBuffer As BaseBuffer
            Private originSnapshot As ITextSnapshot
            Private sourceToken As Object
            Private bufferLength As Integer
            Private changes As List(Of TextChange)
            Private applied As Boolean
            Private canceled As Boolean

            Public ReadOnly Property Snapshot As ITextSnapshot Implements ITextEdit.Snapshot
                Get
                    Return originSnapshot
                End Get
            End Property

            Public Sub New(baseBuffer As BaseBuffer, originSnapshot As ITextSnapshot, sourceToken As Object)
                _baseBuffer = baseBuffer
                Me.originSnapshot = originSnapshot
                Me.sourceToken = sourceToken
                bufferLength = originSnapshot.Length
                changes = New List(Of TextChange)
            End Sub

            Public Function CanInsert(position As Integer) As Boolean Implements ITextEdit.CanInsert
                If position < 0 OrElse position > bufferLength Then
                    Throw New ArgumentOutOfRangeException("position")
                End If

                If canceled Then
                    Throw New InvalidOperationException("Edit has ben canceled")
                End If

                Dim rm = _baseBuffer._readOnlyRegionManager
                If rm Is Nothing Then Return True

                Dim spans = rm.GetReadOnlySpans()
                If spans.Count = 0 Then Return True

                Return Not spans.Exists(
                    Function(s As ReadOnlySpan) Not s.IsInsertAllowed(position)
                )
            End Function

            Public Function CanDeleteOrReplace(span1 As Span) As Boolean Implements ITextEdit.CanDeleteOrReplace
                If span1.Start < 0 OrElse span1.Start > bufferLength Then
                    Throw New ArgumentOutOfRangeException("span")
                End If

                If span1.End > bufferLength Then
                    Throw New ArgumentOutOfRangeException("span")
                End If

                If canceled Then
                    Throw New InvalidOperationException("Edit has ben canceled")
                End If

                If _baseBuffer.ReadOnlyRegionManager Is Nothing Then
                    Return True
                End If

                If _baseBuffer._readOnlyRegionManager.GetReadOnlySpans().Count = 0 Then
                    Return True
                End If

                Return Not _baseBuffer._readOnlyRegionManager.GetReadOnlySpans().Exists(Function(s As ReadOnlySpan) Not s.IsReplaceAllowed(span1))
            End Function

            Public Function Insert(position As Integer, text As String) As Boolean Implements ITextEdit.Insert
                If canceled Then
                    Throw New InvalidOperationException("Edit has ben canceled")
                End If

                If applied Then
                    Throw New InvalidOperationException("Attempt to reuse already applied TextEdit")
                End If

                If position < 0 OrElse position > bufferLength Then
                    Throw New ArgumentOutOfRangeException("position")
                End If

                If text Is Nothing Then
                    Throw New ArgumentNullException("text")
                End If

                If Not CanInsert(position) Then Return False

                For n = changes.Count - 1 To 0 Step -1
                    Dim change = changes(n)
                    If change.Position = position Then
                        changes(n) = New TextChange(position, "", change.NewText & text)
                        Return True
                    End If
                Next

                changes.Add(New TextChange(position, "", text))
                Return True
            End Function

            Public Function Insert(position As Integer, characterBuffer As Char(), startIndex As Integer, length1 As Integer) As Boolean Implements ITextEdit.Insert
                If canceled Then
                    Throw New InvalidOperationException("Edit has ben canceled")
                End If

                If applied Then
                    Throw New InvalidOperationException("Attempt to reuse already applied TextEdit")
                End If

                If position < 0 OrElse position > bufferLength Then
                    Throw New ArgumentOutOfRangeException("position")
                End If

                If characterBuffer Is Nothing Then
                    Throw New ArgumentNullException("characterBuffer")
                End If

                If startIndex < 0 OrElse startIndex > characterBuffer.Length Then
                    Throw New ArgumentOutOfRangeException("startIndex")
                End If

                If length1 < 0 OrElse startIndex + length1 > characterBuffer.Length Then
                    Throw New ArgumentOutOfRangeException("length")
                End If

                If Not CanInsert(position) Then Return False

                changes.Add(New TextChange(position, "", New String(characterBuffer, startIndex, length1)))
                Return True
            End Function

            Public Function Replace(startPosition As Integer, charsToReplace As Integer, replaceWith As String) As Boolean Implements ITextEdit.Replace
                If canceled Then
                    Throw New InvalidOperationException("Edit has ben canceled")
                End If

                If applied Then
                    Throw New InvalidOperationException("Attempt to reuse already applied TextEdit")
                End If

                If startPosition < 0 OrElse startPosition > bufferLength Then
                    Throw New ArgumentOutOfRangeException("startPosition")
                End If

                If charsToReplace < 0 OrElse startPosition + charsToReplace > bufferLength Then
                    Throw New ArgumentOutOfRangeException("charsToReplace")
                End If

                If replaceWith Is Nothing Then
                    replaceWith = ""
                End If

                If Not CanDeleteOrReplace(New Span(startPosition, charsToReplace)) Then
                    Return False
                End If

                changes.Add(New TextChange(startPosition, originSnapshot.GetText(startPosition, charsToReplace), replaceWith))
                Return True
            End Function

            Public Function Replace(replaceSpan As Span, replaceWith As String) As Boolean Implements ITextEdit.Replace
                If canceled Then
                    Throw New InvalidOperationException("Edit has ben canceled")
                End If

                If applied Then
                    Throw New InvalidOperationException("Attempt to reuse already applied TextEdit")
                End If

                If replaceSpan.End > bufferLength Then
                    Throw New ArgumentOutOfRangeException("replaceSpan")
                End If

                If replaceWith Is Nothing Then
                    Throw New ArgumentNullException("replaceWith")
                End If

                If Not CanDeleteOrReplace(replaceSpan) Then Return False

                changes.Add(New TextChange(replaceSpan.Start, originSnapshot.GetText(replaceSpan), replaceWith))
                Return True
            End Function

            Public Function Delete(startPosition As Integer, charsToDelete As Integer) As Boolean Implements ITextEdit.Delete
                If canceled Then
                    Throw New InvalidOperationException("Edit has ben canceled")
                End If

                If applied Then
                    Throw New InvalidOperationException("Attempt to reuse already applied TextEdit")
                End If

                If startPosition < 0 OrElse startPosition > bufferLength Then
                    Throw New ArgumentOutOfRangeException("startPosition")
                End If

                If charsToDelete < 0 OrElse startPosition + charsToDelete > bufferLength Then
                    Throw New ArgumentOutOfRangeException("charsToDelete")
                End If

                If Not CanDeleteOrReplace(New Span(startPosition, charsToDelete)) Then
                    Return False
                End If

                changes.Add(New TextChange(startPosition, originSnapshot.GetText(startPosition, charsToDelete), ""))
                Return True
            End Function

            Public Function Delete(deleteSpan As Span) As Boolean Implements ITextEdit.Delete
                If canceled Then
                    Throw New InvalidOperationException("Edit has ben canceled")
                End If

                If applied Then
                    Throw New InvalidOperationException("Attempt to reuse already applied TextEdit")
                End If

                If deleteSpan.End > bufferLength Then
                    Throw New ArgumentOutOfRangeException("deleteSpan")
                End If

                If Not CanDeleteOrReplace(deleteSpan) Then Return False

                changes.Add(New TextChange(deleteSpan.Start, originSnapshot.GetText(deleteSpan), ""))
                Return True
            End Function

            Public Function Apply() As ITextSnapshot Implements ITextEdit.Apply
                If canceled Then
                    Throw New InvalidOperationException("Edit has ben canceled")
                End If

                If applied Then
                    Throw New InvalidOperationException("Attempt to reuse already applied TextEdit")
                End If

                applied = True
                If changes.Count > 0 Then
                    Return _baseBuffer.FinishChangeApplication(
                          _baseBuffer.ApplyChanges(changes, sourceToken),
                          originSnapshot,
                          sourceToken
                    )
                End If

                Return _baseBuffer._CurrentSnapshot
            End Function

            Public Sub Dispose() Implements IDisposable.Dispose
                If Not applied Then canceled = True
            End Sub

            Public Sub Cancel() Implements ITextEdit.Cancel
                If applied Then
                    Throw New InvalidOperationException("Cannot cancel an applied edit")
                End If
                canceled = True
            End Sub

            Public Sub RemoveLastChange() Implements ITextEdit.RemoveLastChange
                If changes.Count > 0 Then
                    changes.RemoveAt(changes.Count - 1)
                End If
            End Sub
        End Class

        Private _contentType As String
        Private _properties As PropertyCollection
        Private _syncLock As New Object
        Private _readOnlyRegionManager As ReadOnlySpanManagerImpl

        Public Property ContentType As String Implements ITextBuffer.ContentType
            Get
                Return _contentType
            End Get

            Set(value As String)
                _contentType = value
                RaiseEvent ContentTypeChanged(Me, EventArgs.Empty)
            End Set
        End Property

        Public ReadOnly Property Properties As PropertyCollection Implements ITextBuffer.Properties
            Get
                If _properties Is Nothing Then
                    SyncLock _syncLock
                        If _properties Is Nothing Then
                            _properties = New PropertyCollection
                        End If
                    End SyncLock
                End If

                Return _properties
            End Get
        End Property

        Public Property CurrentSnapshot As ITextSnapshot Implements ITextBuffer.CurrentSnapshot

        Public ReadOnly Property ReadOnlyRegionManager As ReadOnlyRegionManager Implements ITextBuffer.ReadOnlyRegionManager
            Get
                If _readOnlyRegionManager Is Nothing Then
                    _readOnlyRegionManager = New ReadOnlySpanManagerImpl(Me)
                End If

                Return _readOnlyRegionManager
            End Get
        End Property

        Private _currentVersion As TextVersionImpl

        Protected ReadOnly Property CurrentVersion As TextVersion
            Get
                Return _currentVersion
            End Get
        End Property

        Public Event ContentTypeChanged As EventHandler Implements ITextBuffer.ContentTypeChanged

        Protected Event ChangedHighPriority As EventHandler(Of TextChangedEventArgs)

        Protected Event ChangedLowPriority As EventHandler(Of TextChangedEventArgs)

        Public Custom Event Changed As EventHandler(Of TextChangedEventArgs) Implements ITextBuffer.Changed
            AddHandler(value As EventHandler(Of TextChangedEventArgs))
                If TypeOf value.Target Is ITextBuffer Then
                    AddHandler ChangedHighPriority, value
                Else
                    AddHandler ChangedLowPriority, value
                End If
            End AddHandler

            RemoveHandler(value As EventHandler(Of TextChangedEventArgs))
                If TypeOf value.Target Is ITextBuffer Then
                    RemoveHandler ChangedHighPriority, value
                Else
                    RemoveHandler ChangedLowPriority, value
                End If
            End RemoveHandler

            RaiseEvent(sender As Object, e As TextChangedEventArgs)
                RaiseEvent ChangedHighPriority(Me, e)
                RaiseEvent ChangedLowPriority(Me, e)
            End RaiseEvent
        End Event

        Protected Sub New(contentType As String)
            Me.ContentType = contentType
            _currentVersion = New TextVersionImpl(Me, 0)
        End Sub

        Public Function CreateEdit(sourceToken As Object) As ITextEdit Implements ITextBuffer.CreateEdit
            Return New Edit(Me, _CurrentSnapshot, sourceToken)
        End Function

        Public Function CreateEdit() As ITextEdit Implements ITextBuffer.CreateEdit
            Return CreateEdit(Nothing)
        End Function

        Public Function Insert(position As Integer, text As String) As ITextSnapshot Implements ITextBuffer.Insert
            Using textEdit = CreateEdit()
                textEdit.Insert(position, text)
                Return textEdit.Apply()
            End Using
        End Function

        Public Function Delete(deleteSpan As Span) As ITextSnapshot Implements ITextBuffer.Delete
            Using textEdit = CreateEdit()
                textEdit.Delete(deleteSpan)
                Return textEdit.Apply()
            End Using
        End Function

        Public Function Replace(replaceSpan As Span, replaceWith As String) As ITextSnapshot Implements ITextBuffer.Replace
            Using textEdit = CreateEdit()
                textEdit.Replace(replaceSpan, replaceWith)
                Return textEdit.Apply()
            End Using
        End Function

        Protected MustOverride Function TakeSnapshot() As ITextSnapshot

        Protected Function FinishChangeApplication(normalizedChanges As NormalizedTextChangeCollection, beforeSnapshot As ITextSnapshot, sourceToken As Object) As ITextSnapshot
            SetCurrentVersionAndSnapshot(normalizedChanges)
            RaiseChangedEvent(beforeSnapshot, _CurrentSnapshot, normalizedChanges, sourceToken)
            Return _CurrentSnapshot
        End Function

        Protected Sub SetCurrentVersionAndSnapshot(normalizedChanges As NormalizedTextChangeCollection)
            _currentVersion = _currentVersion.CreateNext(normalizedChanges)
            _CurrentSnapshot = TakeSnapshot()
        End Sub

        Protected MustOverride Function ApplyChanges(changes As List(Of TextChange), sourceToken As Object) As NormalizedTextChangeCollection

        Protected Sub RaiseChangedEvent(beforeSnapshot As ITextSnapshot, afterSnapshot As ITextSnapshot, normalizedChanges As NormalizedTextChangeCollection, sourceToken As Object)
            Dim e As New TextChangedEventArgs(beforeSnapshot, afterSnapshot, normalizedChanges, sourceToken)
            RaiseEvent ChangedHighPriority(Me, e)
            RaiseEvent ChangedLowPriority(Me, e)
        End Sub
    End Class
End Namespace
