Imports System.Collections
Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports System.Text
Imports System.Windows.Input
Imports System.Windows.Threading
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations
Imports System.Runtime.InteropServices

Namespace SuperClassifier

    <ComponentOptions(ComponentDiscoveryMode:=ComponentDiscoveryMode.Always)>
    Public NotInheritable Class SmartIndentor
        Inherits KeyboardFilter

        Private Enum IndentAction
            KeepSame
            Indent
            Unindent
        End Enum

        Private Enum IndendationType
            None
            IndentingWord
            UnindentingWord
        End Enum

        Private Class IndentationResolver
            Private isCaseSensitive As Boolean
            Private indentingHash As New Hashtable
            Private unindentingHash As New Hashtable

            Public Sub New(indentingWord As IEnumerable(Of String), unintentingWord As IEnumerable(Of String), isCaseSensitive As Boolean)
                Me.isCaseSensitive = isCaseSensitive
                If Me.isCaseSensitive Then
                    For Each word In indentingWord
                        indentingHash.Add(word, Nothing)
                    Next

                    For Each word In unintentingWord
                        unindentingHash.Add(word, Nothing)
                    Next

                Else
                    For Each word In indentingWord
                        indentingHash.Add(word.ToLowerInvariant(), Nothing)
                    Next

                    For Each word In unintentingWord
                        unindentingHash.Add(word.ToLowerInvariant(), Nothing)
                    Next
                End If
            End Sub

            Public Function IsIndentingWord(word As String) As Boolean
                If Not isCaseSensitive Then
                    word = word.ToLower()
                End If

                If indentingHash.Contains(word) Then
                    Return True
                End If

                Return False
            End Function

            Public Function IsUnindentingWord(word As String) As Boolean
                If Not isCaseSensitive Then
                    word = word.ToLower()
                End If

                If unindentingHash.Contains(word) Then
                    Return True
                End If
                Return False
            End Function
        End Class

        Public Const Name As String = "SmartIndentor"
        Private indentators As New Dictionary(Of String, IndentationResolver)

        <Import>
        Public Property EditorOperationsProvider As IEditorOperationsProvider

        Public Overrides ReadOnly Property IsApplicableToHandledEvents As Boolean = True

        Public Overrides Sub KeyDown(view As IAvalonTextView, args As KeyEventArgs)
            If args.Key <> Key.Return Then Return

            Dim tokenizer1 As Tokenizer = SuperClassifierProvider.GetTokenizer(view.TextBuffer)
            If tokenizer1 Is Nothing Then Return

            view.VisualElement.Dispatcher.BeginInvoke(DispatcherPriority.Normal, CType(
                       Function() As Object
                           Dim characterIndex1 As Integer = view.Caret.Position.CharacterIndex
                           Dim currentSnapshot1 As ITextSnapshot = view.TextBuffer.CurrentSnapshot
                           Dim lineNumberFromPosition As Integer = currentSnapshot1.GetLineNumberFromPosition(characterIndex1)

                           If lineNumberFromPosition > 0 Then
                               Dim lineFromLineNumber As ITextSnapshotLine = currentSnapshot1.GetLineFromLineNumber(lineNumberFromPosition - 1)
                               Dim tokensCovering As List(Of Token) = tokenizer1.GetTokensCovering(lineFromLineNumber.Start, lineFromLineNumber.Length)
                               Dim previousToken As Token = GetPreviousTokenAtPosition(tokensCovering, characterIndex1)

                               If previousToken IsNot Nothing Then
                                   Dim resolver As IndentationResolver = GetIndentationResolver(view.TextBuffer.ContentType)

                                   If resolver IsNot Nothing AndAlso lineFromLineNumber.Start <= previousToken.TokenStart Then
                                       view.VisualElement.Dispatcher.BeginInvoke(DispatcherPriority.Normal, CType(
                                                Function() As Object
                                                    If resolver.IsIndentingWord(previousToken.TokenString) Then
                                                        ApplyIndentationIfNecessary(view, IndentAction.Indent)
                                                    Else
                                                        ApplyIndentationIfNecessary(view, IndentAction.KeepSame)
                                                    End If
                                                    Return Nothing
                                                End Function, DispatcherOperationCallback), Nothing)
                                   End If
                               End If
                           End If
                           Return Nothing
                       End Function, DispatcherOperationCallback), Nothing)
        End Sub

        Public Overrides Sub TextInput(view As IAvalonTextView, args As TextCompositionEventArgs)
            Dim tokenizer1 As Tokenizer = SuperClassifierProvider.GetTokenizer(view.TextBuffer)
            If tokenizer1 Is Nothing Then Return

            Dim characterIndex1 As Integer = view.Caret.Position.CharacterIndex
            Dim lineFromPosition As ITextSnapshotLine = view.TextBuffer.CurrentSnapshot.GetLineFromPosition(characterIndex1)
            If lineFromPosition.LineNumber <= 0 Then Return

            Dim lineFromLineNumber As ITextSnapshotLine = view.TextBuffer.CurrentSnapshot.GetLineFromLineNumber(lineFromPosition.LineNumber - 1)
            Dim tokensCovering As List(Of Token) = tokenizer1.GetTokensCovering(lineFromLineNumber.Start, lineFromPosition.End - lineFromLineNumber.Start)
            Dim previousTokenAtPosition As Token = GetPreviousTokenAtPosition(tokensCovering, characterIndex1)
            If previousTokenAtPosition Is Nothing Then Return

            Dim resolver As IndentationResolver = GetIndentationResolver(view.TextBuffer.ContentType)
            If resolver Is Nothing OrElse Not resolver.IsUnindentingWord(previousTokenAtPosition.TokenString) Then
                Return
            End If

            Dim previousToPreviousToken As Token = GetPreviousTokenAtPosition(tokensCovering, previousTokenAtPosition.TokenStart)
            If previousToPreviousToken Is Nothing OrElse previousToPreviousToken.TokenEnd > lineFromLineNumber.EndIncludingLineBreak Then
                Return
            End If

            view.VisualElement.Dispatcher.BeginInvoke(DispatcherPriority.Normal, CType(
                     Function() As Object
                         If resolver.IsIndentingWord(previousToPreviousToken.TokenString) Then
                             ApplyIndentationIfNecessary(view, IndentAction.KeepSame)
                         Else
                             ApplyIndentationIfNecessary(view, IndentAction.Unindent)
                         End If
                         Return Nothing
                     End Function, DispatcherOperationCallback), Nothing)
        End Sub

        Private Sub ApplyIndentationIfNecessary(view As IAvalonTextView, action As IndentAction)
            Dim characterIndex1 As Integer = view.Caret.Position.CharacterIndex
            Dim lineFromPosition As ITextSnapshotLine = view.TextBuffer.CurrentSnapshot.GetLineFromPosition(characterIndex1)
            If lineFromPosition.LineNumber = 0 Then Return

            Dim lineFromLineNumber As ITextSnapshotLine = view.TextBuffer.CurrentSnapshot.GetLineFromLineNumber(lineFromPosition.LineNumber - 1)
            Dim editorOperations As IEditorOperations = EditorOperationsProvider.GetEditorOperations(view)
            Dim spaceCount As Integer
            Dim __ As Integer
            CountSpaces(lineFromLineNumber, editorOperations.TabSize, spaceCount, __)

            Dim spaceCount2 As Integer
            Dim charCount2 As Integer
            CountSpaces(lineFromPosition, editorOperations.TabSize, spaceCount2, charCount2)

            Select Case action
                Case IndentAction.Indent
                    If spaceCount2 <> spaceCount + editorOperations.TabSize Then
                        view.TextBuffer.Replace(New Span(lineFromPosition.Start, charCount2), GetIndentationText(editorOperations.ConvertTabsToSpace, editorOperations.TabSize, spaceCount + editorOperations.TabSize))
                    End If

                Case IndentAction.Unindent
                    If spaceCount2 <> spaceCount - editorOperations.TabSize Then
                        view.TextBuffer.Replace(New Span(lineFromPosition.Start, charCount2), GetIndentationText(editorOperations.ConvertTabsToSpace, editorOperations.TabSize, spaceCount - editorOperations.TabSize))
                    End If

                Case IndentAction.KeepSame
                    If spaceCount2 <> spaceCount Then
                        view.TextBuffer.Replace(New Span(lineFromPosition.Start, charCount2), GetIndentationText(editorOperations.ConvertTabsToSpace, editorOperations.TabSize, spaceCount))
                    End If

            End Select
        End Sub

        Private Function GetIndentationText(isTabsToSpace As Boolean, tabSize1 As Integer, spaceCount As Integer) As String
            Dim value As String = (If(isTabsToSpace, New String(" "c, tabSize1), vbTab))
            Dim stringBuilder1 As New StringBuilder
            Dim i As Integer
            i = 0

            While i + tabSize1 <= spaceCount
                stringBuilder1.Append(value)
                i += tabSize1
            End While

            While i < spaceCount
                stringBuilder1.Append(" "c)
                i += 1
            End While

            Return stringBuilder1.ToString()
        End Function

        Private Sub CountSpaces(line As ITextSnapshotLine, tabSize1 As Integer, <Out> ByRef spaceCount As Integer, <Out> ByRef charCount As Integer)
            spaceCount = 0
            charCount = 0

            For i As Integer = line.Start To line.End - 1
                If line.TextSnapshot(i) = " "c Then
                    spaceCount += 1
                    charCount += 1
                    Continue For
                End If

                If line.TextSnapshot(i) = vbTab Then
                    spaceCount += tabSize1
                    charCount += 1
                    Continue For
                End If
                Exit For
            Next
        End Sub

        Private Function GetIndentationResolver(contentTypeName As String) As IndentationResolver
            Dim value As SuperClassifier.SmartIndentor.IndentationResolver = Nothing
            If Not indentators.TryGetValue(contentTypeName, value) Then
                Dim languageData1 As LanguageData = SuperClassifierProvider.GetLanguageData(contentTypeName)
                If languageData1 IsNot Nothing Then
                    value = New IndentationResolver(languageData1.IndentingWords, languageData1.UnindentingWords, languageData1.IsCaseSensitive)
                End If

                indentators(contentTypeName) = value
            End If

            Return value
        End Function

        Private Function GetPreviousTokenAtPosition(tokenList As IList(Of Token), position1 As Integer) As Token
            For num As Integer = tokenList.Count - 1 To 0 Step -1
                If tokenList(num).TokenStart < position1 Then
                    Return tokenList(num)
                End If
            Next
            Return Nothing
        End Function
    End Class
End Namespace
