Imports System.Windows.Threading
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations
Imports Microsoft.SmallBasic.Completion
Imports System.Runtime.InteropServices
Imports sb = Microsoft.SmallBasic

Namespace Microsoft.SmallBasic.LanguageService
    Public NotInheritable Class CompletionAdornmentProvider
        Implements IAdornmentProvider

        Private textBuffer As ITextBuffer
        Private textView As ITextView
        Private completionHelper As CompletionHelper
        Private adornment As CompletionAdornment
        Private dismissedSpan As ITextSpan
        Private helpUpdateTimer As DispatcherTimer
        Private editorOperations As IEditorOperations
        Private undoHistory As UndoHistory
        Public Event AdornmentsChanged As EventHandler(Of AdornmentsChangedEventArgs) Implements IAdornmentProvider.AdornmentsChanged

        Public Sub New(ByVal textView As ITextView, ByVal editorOperationsProvider As IEditorOperationsProvider, ByVal undoHistoryRegistry As IUndoHistoryRegistry)
            Me.textView = textView
            textBuffer = textView.TextBuffer
            AddHandler textBuffer.Changed, AddressOf OnTextBufferChanged
            editorOperations = editorOperationsProvider.GetEditorOperations(textView)
            undoHistory = undoHistoryRegistry.GetHistory(textView.TextBuffer)
            helpUpdateTimer = New DispatcherTimer(TimeSpan.FromMilliseconds(500.0), DispatcherPriority.ApplicationIdle, AddressOf OnHelpUpdate, Application.Current.Dispatcher)
            helpUpdateTimer.Stop()
            AddHandler Me.textView.LayoutChanged, AddressOf OnLayoutChanged

            If Not textBuffer.Properties.TryGetProperty(GetType(CompletionHelper), completionHelper) Then
                completionHelper = New CompletionHelper()
                textBuffer.Properties.AddProperty(GetType(CompletionHelper), completionHelper)
            End If
        End Sub

        Public Sub CommitItem(ByVal item As CompletionItem)
            If adornment IsNot Nothing Then
                Dim replacementText = item.ReplacementText
                Dim replaceSpan As Span = adornment.ReplaceSpan.GetSpan(textView.TextSnapshot)
                editorOperations.ReplaceText(replaceSpan, replacementText, undoHistory)
                DismissAdornment(force:=True)
                TryCast(textView, Control)?.Focus()
            End If
        End Sub

        Public Sub DismissAdornment(ByVal force As Boolean)
            If adornment Is Nothing Then
                Return
            End If

            dismissedSpan = adornment.Span
            Dim span = dismissedSpan.GetSpan(textView.TextSnapshot)

            If span.Length = 0 AndAlso span.Start > 0 Then
                dismissedSpan = New TextSpan(textView.TextSnapshot, span.Start - 1, 1, SpanTrackingMode.EdgeInclusive)
            End If

            adornment = Nothing

            If AdornmentsChangedEvent IsNot Nothing Then
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, CType(Function()
                                                                                           RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(dismissedSpan))
                                                                                           Return Nothing
                                                                                       End Function, DispatcherOperationCallback), Nothing)
            End If

            If Not force Then
                dismissedSpan = Nothing
            End If
        End Sub

        Public Function GetAdornments(ByVal span As SnapshotSpan) As IList(Of IAdornment) Implements IAdornmentProvider.GetAdornments
            Dim list As List(Of IAdornment) = New List(Of IAdornment)()

            If adornment IsNot Nothing AndAlso adornment.Span.GetSpan(span.Snapshot).OverlapsWith(span) Then
                list.Add(adornment)
            End If

            Return list
        End Function

        Private Sub OnLayoutChanged(ByVal sender As Object, ByVal e As TextViewLayoutChangedEventArgs)
            AddHandler textView.Caret.PositionChanged, AddressOf OnCaretPositionChanged
            RemoveHandler textView.LayoutChanged, AddressOf OnLayoutChanged
        End Sub

        Private Sub OnCaretPositionChanged(ByVal sender As Object, ByVal e As CaretPositionChangedEventArgs)
            If adornment IsNot Nothing Then
                Dim span = adornment.Span.GetSpan(textView.TextSnapshot)
                Dim textInsertionIndex = e.NewPosition.TextInsertionIndex

                If textInsertionIndex < span.Start OrElse textInsertionIndex > span.End Then
                    DismissAdornment(force:=False)
                End If
            Else
                helpUpdateTimer.Start()
            End If
        End Sub

        Private Sub OnHelpUpdate(ByVal sender As Object, ByVal e As EventArgs)
            helpUpdateTimer.Stop()
            Dim textInsertionIndex = textView.Caret.Position.TextInsertionIndex
            Dim lineFromPosition = textView.TextSnapshot.GetLineFromPosition(textInsertionIndex)
            Dim currentToken As TokenInfo
            Dim completionBag = GetCompletionBag(lineFromPosition, textInsertionIndex - lineFromPosition.Start, currentToken)

            If completionBag Is Nothing OrElse completionBag.CompletionItems.Count <= 0 OrElse currentToken.Token = Token.Illegal Then
                Return
            End If

            Dim lineScanner As LineScanner = New LineScanner()
            lineScanner.GetTokenList(lineFromPosition.GetText(), lineFromPosition.LineNumber)

            For Each completionItem In completionBag.CompletionItems

                If String.Equals(completionItem.DisplayName, currentToken.Text, StringComparison.InvariantCultureIgnoreCase) Then
                    UpdateCurrentCompletionItem(New CompletionItemWrapper(completionItem))
                    Exit For
                End If
            Next
        End Sub

        Private Sub OnTextBufferChanged(ByVal sender As Object, ByVal e As Nautilus.Text.TextChangedEventArgs)
            Dim textChange = e.Changes(0)
            Dim newEnd = textChange.NewEnd

            If adornment IsNot Nothing Then
                Dim span = adornment.Span.GetSpan(e.After)

                If span.IsEmpty OrElse newEnd < span.Start OrElse newEnd > span.End Then
                    DismissAdornment(force:=False)
                End If
            ElseIf textChange.NewText.Length = 1 AndAlso (Char.IsLetter(textChange.NewText(0)) OrElse textChange.NewText(0) = "."c) Then
                ShowCompletionAdornment(e.After, newEnd)
            End If
        End Sub

        Public Function GetCompletionBag(ByVal line As ITextSnapshotLine, ByVal column As Integer, <Out> ByRef currentToken As TokenInfo) As CompletionBag
            currentToken = TokenInfo.Illegal
            Dim completionBag = GetMemberCompletionBag(line, column, currentToken)

            If currentToken.Token = Token.StringLiteral OrElse currentToken.Token = Token.Comment Then
                Return completionBag
            End If

            If completionBag Is Nothing OrElse completionBag.CompletionItems.Count = 0 Then
                Dim source As TextBufferReader = New TextBufferReader(line.TextSnapshot)
                completionBag = completionHelper.GetCompletionItems(source, line.LineNumber, column)
            End If

            Return completionBag
        End Function

        Public Sub ShowCompletionAdornment(ByVal snapshot As ITextSnapshot, ByVal caretPosition As Integer)
            Dim lineFromPosition = snapshot.GetLineFromPosition(caretPosition)
            Dim currentToken As TokenInfo
            Dim completionBag = GetCompletionBag(lineFromPosition, caretPosition - lineFromPosition.Start, currentToken)

            If completionBag Is Nothing OrElse completionBag.CompletionItems.Count <= 0 Then
                Return
            End If

            Dim adornmentSpan = GetTextSpanFromToken(lineFromPosition, currentToken)
            Dim textSpan = adornmentSpan

            If textSpan.GetSpan(lineFromPosition.TextSnapshot).IsEmpty AndAlso lineFromPosition.TextSnapshot.Length > 0 Then
                If currentToken.Column = 0 Then
                    adornmentSpan = New TextSpan(lineFromPosition.TextSnapshot, lineFromPosition.Start, Math.Min(currentToken.EndColumn - currentToken.Column + 1, lineFromPosition.TextSnapshot.Length), SpanTrackingMode.EdgeInclusive)
                Else
                    adornmentSpan = New TextSpan(lineFromPosition.TextSnapshot, lineFromPosition.Start + currentToken.Column - 1, Math.Min(currentToken.EndColumn - currentToken.Column + 1, lineFromPosition.TextSnapshot.Length), SpanTrackingMode.EdgeInclusive)
                End If
            End If

            adornment = New CompletionAdornment(Me, completionBag, adornmentSpan, textSpan)

            If AdornmentsChangedEvent IsNot Nothing Then
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, CType(Function()
                                                                                           RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(adornmentSpan))
                                                                                           Return Nothing
                                                                                       End Function, DispatcherOperationCallback), Nothing)
            End If
        End Sub

        Private Function GetMemberCompletionBag(ByVal line As ITextSnapshotLine, ByVal column As Integer, <Out> ByRef currentToken As TokenInfo) As CompletionBag
            currentToken = sb.TokenInfo.Illegal
            Dim lineScanner As New LineScanner()
            Dim tokenList = lineScanner.GetTokenList(line.GetText(), line.LineNumber)
            Dim _tokenInfo = TokenInfo.Illegal
            Dim illegal = TokenInfo.Illegal

            Do
                illegal = _tokenInfo
                _tokenInfo = currentToken
                currentToken = tokenList.Current

            Loop While (column < currentToken.Column OrElse column > currentToken.EndColumn) AndAlso
                                column >= currentToken.Column AndAlso tokenList.MoveNext()


            Dim tokenInfo2 = TokenInfo.Illegal

            If _tokenInfo.Token = Token.Dot Then
                tokenInfo2 = illegal
            ElseIf currentToken.Token = Token.Dot Then
                tokenInfo2 = _tokenInfo
                Dim endColumn = currentToken.EndColumn
                currentToken = TokenInfo.Illegal
                currentToken.Column = endColumn
                currentToken.Text = ""
            End If

            If tokenInfo2.Token <> 0 Then
                Dim value As TypeInfo = Nothing
                Dim emptyCompletionBag As CompletionBag = completionHelper.GetEmptyCompletionBag()

                If emptyCompletionBag.TypeInfoBag.Types.TryGetValue(tokenInfo2.NormalizedText, value) Then
                    CompletionHelper.FillMemberNames(emptyCompletionBag, value)
                Else

                End If

                emptyCompletionBag.CompletionItems.Sort(Function(ByVal ci1, ByVal ci2) ci1.DisplayName.CompareTo(ci2.DisplayName))
                Return emptyCompletionBag
            End If

            Return Nothing
        End Function

        Private Function GetTextSpanFromToken(ByVal line As ITextSnapshotLine, ByVal tokenInfo As TokenInfo) As ITextSpan
            If tokenInfo.Token = Token.Illegal AndAlso tokenInfo.Column = 0 Then
                Return line.TextSnapshot.CreateTextSpan(line.End, 0, SpanTrackingMode.EdgeInclusive)
            End If

            Return line.TextSnapshot.CreateTextSpan(line.Start + tokenInfo.Column, tokenInfo.Text.Length, SpanTrackingMode.EdgeInclusive)
        End Function

        Private Function GetTokenInfo(ByVal tokenEnumerator As TokenEnumerator, ByVal column As Integer) As TokenInfo
            Do
                Dim current = tokenEnumerator.Current

                If current.Column <= column AndAlso current.EndColumn > column Then
                    Return current
                End If
            Loop While tokenEnumerator.MoveNext()

            Return TokenInfo.Illegal
        End Function

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class
End Namespace
