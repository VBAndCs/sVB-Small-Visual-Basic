Imports System.Windows.Threading
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations
Imports Microsoft.SmallBasic.Completion
Imports System.Runtime.InteropServices
Imports System.Collections.ObjectModel

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

        Public Sub New(textView As ITextView, editorOperationsProvider As IEditorOperationsProvider, undoHistoryRegistry As IUndoHistoryRegistry)
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

        Public Sub CommitItem(item As CompletionItem)
            If adornment IsNot Nothing Then
                Dim replacementText = item.ReplacementText
                Dim replaceSpan = adornment.ReplaceSpan.GetSpan(textView.TextSnapshot)
                editorOperations.ReplaceText(replaceSpan, replacementText, undoHistory)
                DismissAdornment(force:=True)
                TryCast(textView, Control)?.Focus()
            End If
        End Sub

        Public Sub DismissAdornment(force As Boolean)
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
                Application.Current.Dispatcher.Invoke(
                    DispatcherPriority.Normal,
                    CType(Function()
                              RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(dismissedSpan))
                              Return Nothing
                          End Function, DispatcherOperationCallback),
                    Nothing)
            End If

            If Not force Then
                dismissedSpan = Nothing
            End If
        End Sub

        Public Function GetAdornments(span As SnapshotSpan) As IList(Of IAdornment) Implements IAdornmentProvider.GetAdornments
            Dim list As List(Of IAdornment) = New List(Of IAdornment)()

            If adornment IsNot Nothing AndAlso adornment.Span.GetSpan(span.Snapshot).OverlapsWith(span) Then
                list.Add(adornment)
            End If

            Return list
        End Function

        Private Sub OnLayoutChanged(sender As Object, e As TextViewLayoutChangedEventArgs)
            AddHandler textView.Caret.PositionChanged, AddressOf OnCaretPositionChanged
            RemoveHandler textView.LayoutChanged, AddressOf OnLayoutChanged
        End Sub

        Private Sub OnCaretPositionChanged(sender As Object, e As CaretPositionChangedEventArgs)
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

        Private Sub OnHelpUpdate(sender As Object, e As EventArgs)
            helpUpdateTimer.Stop()
            Dim textInsertionIndex = textView.Caret.Position.TextInsertionIndex
            Dim lineFromPosition = textView.TextSnapshot.GetLineFromPosition(textInsertionIndex)
            Dim currentToken As TokenInfo
            Dim completionBag = GetCompletionBag(lineFromPosition, textInsertionIndex - lineFromPosition.Start, currentToken)

            If completionBag Is Nothing OrElse completionBag.CompletionItems.Count <= 0 OrElse currentToken.Token = Token.Illegal Then
                Return
            End If

            For Each completionItem In completionBag.CompletionItems
                If String.Equals(completionItem.DisplayName, currentToken.Text, StringComparison.InvariantCultureIgnoreCase) Then
                    UpdateCurrentCompletionItem(New CompletionItemWrapper(completionItem))
                    Exit For
                End If
            Next
        End Sub

        Private Sub OnTextBufferChanged(sender As Object, e As Nautilus.Text.TextChangedEventArgs)
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

        Public Function GetCompletionBag(line As ITextSnapshotLine, column As Integer, <Out> ByRef currentToken As TokenInfo) As CompletionBag
            Dim ControlsInfo As Dictionary(Of String, String) = Nothing
            line.TextSnapshot.TextBuffer.Properties.TryGetProperty("ControlsInfo", ControlsInfo)

            Dim ControlNames As New ObservableCollection(Of String)
            line.TextSnapshot.TextBuffer.Properties.TryGetProperty("ControlNames", ControlNames)

            currentToken = TokenInfo.Illegal
            Dim lineScanner As New LineScanner()
            Dim tokenList = lineScanner.GetTokenList(line.GetText(), line.LineNumber)
            Dim prevToken = TokenInfo.Illegal
            Dim b4PrevToken = TokenInfo.Illegal

            Do
                b4PrevToken = prevToken
                prevToken = currentToken
                currentToken = tokenList.Current

            Loop While (column < currentToken.Column OrElse column > currentToken.EndColumn) AndAlso
                                column >= currentToken.Column AndAlso tokenList.MoveNext()


            Dim identifierToken = TokenInfo.Illegal

            If prevToken.Token = Token.Dot Then
                identifierToken = b4PrevToken
            ElseIf currentToken.Token = Token.Dot Then
                identifierToken = prevToken
                Dim endColumn = currentToken.EndColumn
                currentToken = TokenInfo.Illegal
                currentToken.Column = endColumn
                currentToken.Text = ""
            End If

            Dim newCompletionBag = completionHelper.GetEmptyCompletionBag()
            Dim addGlobals = False

            If identifierToken.Token = TokenInfo.Illegal.Token Then
                addGlobals = True
                If ControlsInfo IsNot Nothing Then
                    Dim txt = currentToken.NormalizedText
                    If txt <> "" Then
                        Dim controls = From name In ControlNames
                                       Where name.ToLower().StartsWith(txt)

                        For Each name In controls
                            newCompletionBag.CompletionItems.Add(New CompletionItem With {
                                           .DisplayName = name,
                                           .ItemType = CompletionItemType.Identifier,
                                           .ReplacementText = name
                                       }
                                )
                        Next
                    End If
                End If
            Else
                Dim value As TypeInfo = Nothing

                If newCompletionBag.TypeInfoBag.Types.TryGetValue(identifierToken.NormalizedText, value) Then
                    CompletionHelper.FillMemberNames(newCompletionBag, value)
                End If

                Dim controlName = identifierToken.NormalizedText
                If ControlsInfo?.ContainsKey(controlName) Then
                    FillMemberNames(newCompletionBag, ControlsInfo(controlName))
                End If
            End If

            If addGlobals OrElse newCompletionBag Is Nothing OrElse newCompletionBag.CompletionItems.Count = 0 Then
                If Not (currentToken.Token = Token.StringLiteral OrElse currentToken.Token = Token.Comment) Then
                    Dim source = New TextBufferReader(line.TextSnapshot)
                    newCompletionBag.CompletionItems.AddRange(completionHelper.GetCompletionItems(source, line.LineNumber, column).CompletionItems)
                End If
            End If

            newCompletionBag.CompletionItems.Sort(Function(ci1, ci2) ci1.DisplayName.CompareTo(ci2.DisplayName))
            Return newCompletionBag
        End Function

        Public Sub ShowCompletionAdornment(snapshot As ITextSnapshot, caretPosition As Integer)
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

        Private Function GetTextSpanFromToken(line As ITextSnapshotLine, tokenInfo As TokenInfo) As ITextSpan
            If tokenInfo.Token = Token.Illegal AndAlso tokenInfo.Column = 0 Then
                Return line.TextSnapshot.CreateTextSpan(line.End, 0, SpanTrackingMode.EdgeInclusive)
            End If

            Return line.TextSnapshot.CreateTextSpan(line.Start + tokenInfo.Column, tokenInfo.Text.Length, SpanTrackingMode.EdgeInclusive)
        End Function

        Private Function GetTokenInfo(tokenEnumerator As TokenEnumerator, column As Integer) As TokenInfo
            Do
                Dim current = tokenEnumerator.Current

                If current.Column <= column AndAlso current.EndColumn > column Then
                    Return current
                End If
            Loop While tokenEnumerator.MoveNext()

            Return TokenInfo.Illegal
        End Function



        Private Shared CompletionItems As New Dictionary(Of String, List(Of CompletionItem))

        Shared Sub New()
            AddCompletionList(GetType(WinForms.Form))
            AddCompletionList(GetType(WinForms.Control))
            AddCompletionList(GetType(WinForms.TextBox))
            AddCompletionList(GetType(WinForms.Label))
            AddCompletionList(GetType(WinForms.Button))
            AddCompletionList(GetType(WinForms.ListBox))
        End Sub

        Private Shared Sub AddCompletionList(type As Type)
            Dim compList As New List(Of CompletionItem)

            Dim methods = type.GetMethods(System.Reflection.BindingFlags.Static Or System.Reflection.BindingFlags.Public)
            Dim extensionParams = If(type.Name = "Form", 1, 2)

            For Each methodInfo In methods
                Dim name = ""
                Dim completionItem As New CompletionItem()
                If methodInfo.GetCustomAttributes(GetType(WinForms.ExMethodAttribute), inherit:=False).Count > 0 Then
                    name = methodInfo.Name
                    completionItem.Name = name
                    completionItem.DisplayName = name
                    completionItem.ItemType = CompletionItemType.MethodName
                    If methodInfo.GetParameters().Length > extensionParams Then
                        completionItem.ReplacementText = name & "("
                    Else
                        completionItem.ReplacementText = name & "()"
                    End If
                    completionItem.MemberInfo = methodInfo
                    compList.Add(completionItem)

                ElseIf methodInfo.Name.ToLower().StartsWith("get") AndAlso methodInfo.GetCustomAttributes(GetType(WinForms.ExPropertyAttribute), inherit:=False).Count > 0 Then
                    name = methodInfo.Name.Substring(3)
                    completionItem.Name = name
                    completionItem.DisplayName = name
                    completionItem.ItemType = CompletionItemType.PropertyName
                    completionItem.ReplacementText = name
                    completionItem.MemberInfo = methodInfo
                    compList.Add(completionItem)
                End If
            Next

            Dim events = type.GetEvents(System.Reflection.BindingFlags.Static Or System.Reflection.BindingFlags.Public)
            For Each eventInfo In events
                If eventInfo.EventHandlerType Is GetType(Library.SmallBasicCallback) Then
                    Dim name = eventInfo.Name
                    compList.Add(New CompletionItem() With {
                    .Name = name,
                    .DisplayName = name,
                    .ItemType = CompletionItemType.EventName,
                    .ReplacementText = name,
                    .MemberInfo = eventInfo
                })
                End If
            Next

            CompletionItems.Add(type.Name, compList)

        End Sub

        Private Sub FillMemberNames(completionBag As CompletionBag, moduleName As String)
            Dim controlModule = NameOf(WinForms.Control)
            completionBag.CompletionItems.AddRange(CompletionItems(controlModule))

            If moduleName <> controlModule Then
                completionBag.CompletionItems.AddRange(CompletionItems(moduleName))
            End If

        End Sub
    End Class
End Namespace
