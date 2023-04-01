Imports System.Globalization
Imports System.Windows.Media.Animation
Imports System.Windows.Threading
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Library

Namespace Microsoft.SmallVisualBasic.LanguageService
    Partial Public Class CompletionSurface
        Inherits Canvas
        Implements IAdornmentSurface

        Private textView As IAvalonTextView
        Private filteredCompletionItems As List(Of CompletionItemWrapper)
        Friend IsCommitting As Boolean
        Public ReadOnly Property Adornment As CompletionAdornment
        Public ReadOnly Property CanNegotiateSurfaceSpace As Boolean = False Implements IAdornmentSurface.CanNegotiateSurfaceSpace
        Public ReadOnly Property SurfacePosition As SurfacePosition = SurfacePosition.AboveText Implements IAdornmentSurface.SurfacePosition
        Public ReadOnly Property SurfacePanel As Panel = Me Implements IAdornmentSurface.SurfacePanel

        Public ReadOnly Property IsAdornmentVisible As Boolean
            Get
                Return CompletionPopup.IsOpen AndAlso Not IsCommitting
            End Get
        End Property


        Public ReadOnly Property TextFlowDirection As FlowDirection
            Get
                Dim currentUICulture = CultureInfo.CurrentUICulture

                If Not currentUICulture.TextInfo.IsRightToLeft Then
                    Return FlowDirection.LeftToRight
                End If

                Return FlowDirection.RightToLeft
            End Get
        End Property

        Public Sub New(textView As IAvalonTextView)
            Me.textView = textView
            AddHandler Me.textView.TextBuffer.Changed, AddressOf OnTextChanged
            InitializeComponent()
            AddHandler Me.CompletionPopup.MouseDown, Sub(sender As Object, e As MouseButtonEventArgs) e.Handled = True
        End Sub

        Public Sub AddAdornment(adornment As IAdornment) Implements IAdornmentSurface.AddAdornment
            _Adornment = TryCast(adornment, CompletionAdornment)
            Dispatcher.BeginInvoke(
                    DispatcherPriority.Normal,
                    CType(Function()
                              DisplayListBox()
                              Return Nothing
                          End Function, DispatcherOperationCallback),
                    Nothing
            )
        End Sub

        Public Sub RemoveAdornment(adornment As IAdornment) Implements IAdornmentSurface.RemoveAdornment
            _Adornment = Nothing
            CompletionPopup.IsOpen = False
        End Sub

        Public Function GetSpaceNegotiations(textLine As ITextLine) As IList(Of SpaceNegotiation) Implements IAdornmentSurface.GetSpaceNegotiations
            Return Nothing
        End Function

        Public Function GetAdornmentsGeometry() As Geometry Implements IAdornmentSurface.GetAdornmentsGeometry
            Return Nothing
        End Function

        Public Sub FadeCompletionList()
            Dim animation As New DoubleAnimation(0.3, New Duration(TimeSpan.FromMilliseconds(300.0)))
            popupContent.BeginAnimation(OpacityProperty, animation)
        End Sub

        Public ReadOnly Property IsFaded() As Boolean
            Get
                Return System.Math.Abs(popupContent.Opacity - 1.0) > Double.Epsilon
            End Get
        End Property

        Public Sub UnfadeCompletionList()
            Dim animation As New DoubleAnimation(1.0, New Duration(TimeSpan.FromMilliseconds(300.0)))
            popupContent.BeginAnimation(OpacityProperty, animation)
        End Sub

        Private Sub OnCompletionClosed(sender As Object, e As EventArgs)
            If _Adornment IsNot Nothing Then
                _Adornment.Dismiss(force:=False)
            End If
        End Sub

        Private Sub CompletionPopup_MouseWheel(sender As Object, e As MouseWheelEventArgs)
            If CompletionListBox IsNot Nothing Then
                If e.Delta < 0 Then
                    CompletionListBox.MoveDown()
                Else
                    CompletionListBox.MoveUp()
                End If

                e.Handled = True
            End If
        End Sub

        Private Sub DisplayListBox()
            If _Adornment IsNot Nothing Then
                Dim span = _Adornment.ReplaceSpan.GetSpan(textView.TextSnapshot)
                Dim textLine = textView.FormattedTextLines.GetTextLineContainingPosition(span.Start)

                If textLine IsNot Nothing Then
                    Dim characterBounds = textLine.GetCharacterBounds(span.Start)
                    Me.CompletionPopup.VerticalOffset = characterBounds.Bottom
                    CompletionPopup.HorizontalOffset = characterBounds.Left
                    FilterItems()
                End If
            End If
        End Sub

        Private Sub FilterItems()
            If CompletionPopup.IsOpen AndAlso _Adornment.AdornmentProvider.adornment Is Nothing Then
                ' Sometimes, the adornmentProvider sets is adornment to nothing, but the adornment still open,
                ' So, we must set the AdornmentProvider.adornment again.
                _Adornment.AdornmentProvider.adornment = _Adornment
            End If

            Dim inputText = GetInputText()
            BuildFilteredCompletionList(inputText, _Adornment.CompletionBag)

            If filteredCompletionItems.Count > 0 Then
                CompletionListBox.ItemsSource = filteredCompletionItems
                CompletionPopup.IsOpen = True
                CompletionListBox.SelectedIndex = GetSelectedItemIndex(inputText)
                UnfadeCompletionList()

            ElseIf CompletionPopup.IsOpen Then
                If _Adornment.CompletionBag.SelectEspecialItem = "" Then
                    FadeCompletionList()
                Else
                    Dim bag = _Adornment.AdornmentProvider.GetOriginalBag()
                    If bag Is Nothing Then
                        _Adornment.Dismiss(True)
                    Else
                        bag.IsBackSpace = _Adornment.CompletionBag.IsBackSpace
                        BuildFilteredCompletionList(inputText, bag)
                        If filteredCompletionItems.Count > 0 Then
                            CompletionListBox.ItemsSource = filteredCompletionItems
                            CompletionPopup.IsOpen = True
                            CompletionListBox.SelectedIndex = GetSelectedItemIndex(inputText)
                            UnfadeCompletionList()
                        Else
                            FadeCompletionList()
                        End If
                    End If
                End If

            Else
                _Adornment?.Dismiss(True)
            End If
        End Sub

        Private Sub OnTextChanged(sender As Object, e As Nautilus.Text.TextChangedEventArgs)
            If IsAdornmentVisible Then
                Dim c = e.Changes(0)
                If c.Delta = -1 Then
                    Dim x = c.OldText
                    Dim bag = _Adornment.CompletionBag
                    bag.IsBackSpace = (x = "(" OrElse x = ",")
                End If
                FilterItems()
            End If
        End Sub

        Private Sub BuildFilteredCompletionList(inputText As String, bag As CompletionBag)
            filteredCompletionItems = New List(Of CompletionItemWrapper)
            Dim items As New List(Of CompletionItemWrapper)

            For Each item In bag.CompletionItems
                If CanAddItem(item, bag.IsFirstToken, inputText) Then
                    Dim itemWrapper = New CompletionItemWrapper(item, bag)
                    items.Add(itemWrapper)
                    filteredCompletionItems.Add(itemWrapper)
                End If
            Next

            If items.Count = 0 Then Return

            If items.Count = 1 Then
                If bag.CtrlSpace Then
                    bag.CtrlSpace = False
                    _Adornment.AdornmentProvider.CommitItem(items(0))
                    filteredCompletionItems.Clear()
                    Return

                ElseIf bag.IsBackSpace OrElse items(0).Display.ToLower() = inputText Then
                    filteredCompletionItems.Clear()
                    _Adornment?.Dismiss(True)
                    Return
                End If
            End If
            ' Completion list must contain 7 items at least
            ' Otherwise repeat the items to fill the gaps
            Do
                Dim n = filteredCompletionItems.Count
                If n > 7 Then Exit Do

                For Each item In items
                    filteredCompletionItems.Add(item)
                Next
            Loop
        End Sub

        Private Function GetUniqueItems(completionItems As CompletionItem()) As List(Of CompletionItem)
            Dim dictionary As Dictionary(Of String, CompletionItem) = New Dictionary(Of String, CompletionItem)()

            For Each completionItem In completionItems
                If Not dictionary.ContainsKey(completionItem.Key) Then
                    dictionary(completionItem.Key) = completionItem
                End If
            Next

            Dim list As List(Of CompletionItem) = New List(Of CompletionItem)(dictionary.Values)
            list.Sort(New Comparison(Of CompletionItem)(AddressOf CompletionItemComparer))
            Return list
        End Function

        Private Function CanAddItem(
                         item As CompletionItem,
                         isFirstToken As Boolean,
                         inputText As String
                    ) As Boolean

            Dim displayName = item.DisplayName

            If displayName.StartsWith("_") Then
                Select Case item.ItemType
                    Case CompletionItemType.LocalVariable,
                             CompletionItemType.GlobalVariable,
                             CompletionItemType.SubroutineName

                    Case Else
                        Return False
                End Select
            End If

            Select Case displayName
                Case "GetHashCode", "ToString", "Equals", "GetType"
                    Return False
            End Select

            If Not isFirstToken Then
                Select Case displayName
                    Case "If", "ElseIf", "Else", "EndIf",
                             "For", "ForEach", "Next", "EndFor",
                             "While", "Wend", "EndWhile",
                             "ExitLoop", "ContinueLoop", "GoTo",
                             "Function", "EndFunction", "Sub", "EndSub", "Return"
                        Return False
                End Select
            End If

            If item.MemberInfo IsNot Nothing AndAlso
                    item.MemberInfo.Name = GetType(Primitive).Name Then Return False

            If inputText = "" Then Return True

            Dim words = GetSubWords(displayName)
            inputText = inputText.ToLowerInvariant()
            For Each word In words
                If word.StartsWith(inputText) Then Return True
            Next

            Return words.Count > 1 AndAlso displayName.ToLowerInvariant().StartsWith(inputText)
        End Function

        Private Function CompletionItemComparer(item1 As CompletionItem, item2 As CompletionItem) As Integer
            Return item1.DisplayName.CompareTo(item2.DisplayName)
        End Function

        Private Function ItemWrapperComparer(item1 As CompletionItemWrapper, item2 As CompletionItemWrapper) As Integer
            Return item1.Display.CompareTo(item2.Display)
        End Function

        Function GetInputText() As String
            Dim replaceSpan = _Adornment.ReplaceSpan
            Dim text = replaceSpan.GetText(textView.TextSnapshot)
            If text.Length = 0 Then
                text = CType(textView, AvalonTextView).Editor.EditorOperations.GetCurrentWord
                If text.Trim.Length = 0 OrElse text.StartsWith(".") OrElse text.StartsWith("!") Then Return ""
            End If

            Dim tokens = LineScanner.GetTokens(text, 0)
            If tokens.Count = 0 Then Return ""
            Dim token = tokens.Last

            If token.Type = TokenType.StringLiteral Then
                Return token.LCaseText.Trim("""")

            ElseIf token.ParseType = ParseType.Operator Then
                Select Case token.Type
                    Case TokenType.Or, TokenType.And, TokenType.RightBracket,
                             TokenType.RightCurlyBracket, TokenType.RightParens

                    Case Else
                        Return ""
                End Select
            End If
            Return token.LCaseText
        End Function

        Private Function GetSelectedItemIndex(text As String) As Integer
            Dim textLength = text.Length
            Dim originalText = text
            Dim key = ""

            If textLength < 2 Then
                Dim firstItem = filteredCompletionItems(0).CompletionItem
                key = firstItem.GetHistoryKey(text)
                If key <> "" Then
                    Dim properties = textView.TextBuffer.Properties
                    Dim controls = properties.GetProperty(Of Dictionary(Of String, String))("ControlsInfo")
                    If controls?.ContainsKey(key) Then key = controls(key).ToLower()

LineRetry:
                    Dim word = GetWord(key)
                    If text = "" OrElse word.StartsWith(text) Then
                        text = word
                        textLength = text.Length
                    End If
                End If
            End If

            If text.Trim = "" Then Return 0

            Dim maxMatchLength = 0
            Dim wordsList As New List(Of List(Of String))

            For i = 0 To filteredCompletionItems.Count - 1
                Dim displayName = filteredCompletionItems(i).Name
                Dim matchLength = GetMatchLength(displayName.ToLowerInvariant(), text)

                If matchLength = textLength Then Return i

                If matchLength > maxMatchLength Then
                    maxMatchLength = matchLength
                End If

                Dim words = GetSubWords(displayName)
                wordsList.Add(words)
            Next

            For i = 0 To wordsList.Count - 1
                Dim words = wordsList(i)

                For j = 1 To words.Count - 1
                    Dim matchLength = GetMatchLength(words(j), text)
                    If matchLength = textLength Then Return i

                    If matchLength > maxMatchLength Then
                        maxMatchLength = matchLength
                    End If
                Next
            Next

            If originalText = "" OrElse originalText = text Then Return 0
            text = originalText
            textLength = text.Length
            If textLength < 2 Then
                key = "__" & text(0)
            End If
            GoTo LineRetry

        End Function

        Private Function GetWord(key As String) As String
            If CompletionHelper.History.ContainsKey(key) Then
                Return CompletionHelper.History(key).ToLower()
            End If

            If Not key.StartsWith("_") OrElse key.StartsWith("_t_") Then Return ""

            key = "_t" & key
            If CompletionHelper.History.ContainsKey(key) Then
                Return CompletionHelper.History(key).ToLower()
            End If

            Return GetBestItem(key.Last)

        End Function

        Private Function GetBestItem(key As String) As String

            For Each item In filteredCompletionItems
                Dim name = item.Display.ToLower()
                If item.SymbolType = SymbolType.Control AndAlso
                   name.StartsWith(key) Then Return name
            Next

            For Each item In filteredCompletionItems
                Dim name = item.Display.ToLower()
                If item.SymbolType = SymbolType.GlobalVariable AndAlso
                   name.StartsWith(key) Then Return name
            Next

            For Each item In filteredCompletionItems
                Dim name = item.Display.ToLower()
                If item.SymbolType = SymbolType.LocalVariable AndAlso
                   name.StartsWith(key) Then Return name
            Next

            For Each item In filteredCompletionItems
                Dim name = item.Display.ToLower()
                If item.SymbolType = SymbolType.Subroutine AndAlso
                   name.StartsWith(key) Then Return name
            Next

            key = UCase(key(0)) & If(key.Length > 1, key.Substring(1), "")

            For Each item In filteredCompletionItems
                Dim name = item.Display
                If item.SymbolType = SymbolType.Control AndAlso
                   name.Contains(key) Then Return name
            Next

            For Each item In filteredCompletionItems
                Dim name = item.Display
                If item.SymbolType = SymbolType.GlobalVariable AndAlso
                   name.Contains(key) Then Return name
            Next

            For Each item In filteredCompletionItems
                Dim name = item.Display
                If item.SymbolType = SymbolType.LocalVariable AndAlso
                   name.Contains(key) Then Return name
            Next

            For Each item In filteredCompletionItems
                Dim name = item.Display
                If item.SymbolType = SymbolType.Subroutine AndAlso
                   name.Contains(key) Then Return name
            Next

            key = key.ToLower()

            For Each item In filteredCompletionItems
                Dim name = item.Display.ToLower()
                If name.StartsWith(key) Then Return name
            Next

            Return ""
        End Function

        Private Function GetSubWords(text As String) As List(Of String)
            Dim words As New List(Of String)
            Dim start = 0
            For i = 1 To text.Length - 1
                If Char.IsUpper(text(i)) Then
                    Dim word = text.Substring(start, i - start).ToLowerInvariant()
                    words.Add(word)
                    start = i
                End If
            Next
            words.Add(text.Substring(start, text.Length - start).ToLowerInvariant())
            Return words
        End Function

        Private Function GetMatchLength(s1 As String, s2 As String) As Integer
            Dim i As Integer
            i = 0
            Dim L = System.Math.Min(s1.Length, s2.Length)

            While i < L AndAlso s1(i) = s2(i)
                i += 1
            End While

            Return i
        End Function

        Private Sub OnCompletionListDoubleClicked(sender As Object, e As MouseButtonEventArgs)
            If _Adornment IsNot Nothing Then
                Dim completionItemWrapper As CompletionItemWrapper = TryCast(Me.CompletionListBox.SelectedItem, CompletionItemWrapper)

                If completionItemWrapper IsNot Nothing Then
                    _Adornment.AdornmentProvider.CommitItem(completionItemWrapper)
                End If
            End If
        End Sub
    End Class
End Namespace
