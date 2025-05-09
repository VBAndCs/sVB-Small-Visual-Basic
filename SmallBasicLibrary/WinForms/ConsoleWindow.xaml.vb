Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Documents
Imports System.Threading.Tasks
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Library.Internal
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Threading

Namespace WinForms
    Partial Public Class ConsoleWindow
        Private Shared _threadTimer As New System.Threading.Timer(AddressOf ResponeToUser)

        Private Shared Sub ResponeToUser(state As Object)
            If Not HandleKey Then
                SmallBasicApplication.Invoke(Sub() DoEvents())
            End If
        End Sub

        Public Sub Write(text As String)
            paragraph.Inlines.Add(New Run(text) With {
                .Foreground = Color.GetBrush(If(TextWindow.ForeColor = Colors.None, TextWindow.ForegroundColor, TextWindow.ForeColor)),
                .Background = Color.GetBrush(If(TextWindow.BackColor = Colors.None, TextWindow.BackgroundColor, TextWindow.BackColor)),
                .FontFamily = New Media.FontFamily(TextWindow.FontName),
                .FontSize = TextWindow._FontSize,
                .FontWeight = If(CBool(TextWindow.FontBold), FontWeights.Bold, FontWeights.Normal),
                .FontStyle = If(CBool(TextWindow.FontItalic), FontStyles.Italic, FontStyles.Normal),
                .TextDecorations = If(CBool(TextWindow.FontUnderlined), TextDecorations.Underline, Nothing)
            })
            ConsoleBox.ScrollToEnd()
            ConsoleBox.CaretPosition = ConsoleBox.Document.ContentEnd
            ConsoleBox.Focus()
            If Me.WindowState = WindowState.Minimized Then
                SystemCommands.RestoreWindow(Me)
            End If
            Me.Activate()
        End Sub

        Public Sub WriteLine(text As String)
            If text <> "" Then Write(text)

            paragraph.Inlines.Add(New LineBreak())
            ConsoleBox.ScrollToEnd()
            ConsoleBox.CaretPosition = ConsoleBox.Document.ContentEnd

            If text <> "" Then
                If Me.WindowState = WindowState.Minimized Then
                    SystemCommands.RestoreWindow(Me)
                End If
                Me.Activate()
            End If
        End Sub

        Friend isWindowClosed As Boolean = False
        Private isKeyPressed As Boolean = False

        Public Function ReadLine() As String
            ShowInputTextBox()
            Do Until InputTextBox.Visibility = Visibility.Collapsed
                DoEvents()
                If isWindowClosed Then Return ""
            Loop
            Return InputTextBox.Text
        End Function

        Private Sub ShowInputTextBox()
            If Me.WindowState = WindowState.Minimized Then
                SystemCommands.RestoreWindow(Me)
            End If
            Me.Activate()
            Dim rect = GetCaretPos()

            With InputTextBox
                .Text = ""
                .Foreground = Color.GetBrush(If(TextWindow.ForeColor = Colors.None, TextWindow.ForegroundColor, TextWindow.ForeColor))
                .Background = Color.GetBrush(If(TextWindow.BackColor = Colors.None, TextWindow.BackgroundColor, TextWindow.BackColor))
                .FontFamily = New Media.FontFamily(TextWindow.FontName)
                .FontSize = TextWindow._FontSize
                .FontWeight = If(CBool(TextWindow.FontBold), FontWeights.Bold, FontWeights.Normal)
                .FontStyle = If(CBool(TextWindow.FontItalic), FontStyles.Italic, FontStyles.Normal)
                .TextDecorations = If(CBool(TextWindow.FontUnderlined), TextDecorations.Underline, Nothing)
                If ConsoleBox.FlowDirection = FlowDirection.RightToLeft Then
                    .Margin = New Thickness(0, rect.Y, rect.X, 0)
                Else
                    .Margin = New Thickness(rect.X, rect.Y, 0, 0)
                End If
                .Visibility = Visibility.Visible
                .Focus()
            End With
        End Sub

        Private Function GetCaretPos() As System.Windows.Rect
            Dim rect As System.Windows.Rect
            Dim caretPosition = ConsoleBox.CaretPosition
            If ConsoleBox.FlowDirection = FlowDirection.RightToLeft Then
                Dim previousPosition = caretPosition.GetPositionAtOffset(-1)
                If previousPosition IsNot Nothing Then
                    Dim precedingChar = (New TextRange(previousPosition, caretPosition)).Text
                    If precedingChar <> "" AndAlso Char.IsDigit(precedingChar(0)) Then
                        'Add space to fix caret pos within the Arabic numbers
                        caretPosition.InsertTextInRun(" ")
                        ConsoleBox.CaretPosition = ConsoleBox.Document.ContentEnd
                        rect = ConsoleBox.CaretPosition.GetCharacterRect(LogicalDirection.Forward)
                        Return rect
                    End If
                End If
            End If

            Return caretPosition.GetCharacterRect(LogicalDirection.Forward)
        End Function

        Dim allowNumbersOnly = False

        Public Function ReadNumber() As Double
            allowNumbersOnly = True
            lastText = ""
            lastSelectionStart = 0
            lastSelectionLength = 0
            Dim n As Double = 0
            Double.TryParse(ReadLine(), n)
            Return n
        End Function

        Friend Sub Clear()
            paragraph.Inlines.Clear()
            Me.Activate()
        End Sub

        Friend Sub ReadKey(intercept As Boolean)
            Throw New NotImplementedException()
        End Sub

        Private Sub InputTextBox_KeyDown(sender As Object, e As KeyEventArgs)
            Select Case e.Key
                Case Key.Enter
                    CommitInputText()
                Case Key.Escape
                    InputTextBox.Text = ""
                Case Key.Space
                    If allowNumbersOnly Then
                        e.Handled = True
                        Beep()
                    End If
            End Select
        End Sub

        Private Sub CommitInputText()
            InputTextBox.Visibility = Visibility.Collapsed
            allowNumbersOnly = False
            WriteLine(InputTextBox.Text)
            ConsoleBox.Focus()
            ConsoleBox.CaretPosition = ConsoleBox.Document.ContentEnd
        End Sub

        Private Sub Input_LostFocus(sender As Object, e As RoutedEventArgs)
            Dim input = CType(sender, System.Windows.Controls.Control)
            If input.Visibility = Visibility.Visible Then
                Dispatcher.BeginInvoke(
                        Sub() input.Focus()
                )
            End If
        End Sub

        Private Sub Window_Closed(sender As Object, e As EventArgs)
            isWindowClosed = True
            TextWindow._windowVisible = False
            _threadTimer.Change(-1, -1)
            If Forms._forms.Count = 0 Then
                Program.End()
            Else
                Program.ActivateWindow()
            End If
        End Sub

        Private Shared HandleKey As Boolean = False

        Friend Sub WaitForAnyKey()
            isKeyPressed = False
            HandleKey = True
            Do
                DoEvents()
                If isKeyPressed OrElse isWindowClosed Then
                    HandleKey = False
                    Return
                End If
            Loop
        End Sub

        Private Sub Window_PreviewKeyDown(sender As Object, e As KeyEventArgs)
            isKeyPressed = True
            e.Handled = HandleKey
        End Sub

        Dim _readOneKey As Boolean = False

        Friend Function ReadKey() As Char
            _readOneKey = True
            InputTextBox.MaxLength = 1
            Dim c = ReadLine()
            _readOneKey = False
            InputTextBox.MaxLength = Integer.MaxValue
            Return c
        End Function

        Dim lastText As String = ""
        Dim lastSelectionStart As Integer = 0
        Dim lastSelectionLength As Integer = 0
        Dim exitTextChange As Boolean = False

        Private Sub InputTextBox_TextChanged(sender As Object, e As System.Windows.Controls.TextChangedEventArgs)
            If exitTextChange Then Return

            If _readOneKey Then
                CommitInputText()
            ElseIf allowNumbersOnly Then
                Dim n = InputTextBox.Text
                If n = "" OrElse n = "-" OrElse n = "." OrElse Double.TryParse(n, Nothing) Then
                    lastText = InputTextBox.Text
                Else
                    Beep()
                    exitTextChange = True
                    InputTextBox.Text = lastText
                    exitTextChange = False
                    InputTextBox.SelectionStart = lastSelectionStart
                    InputTextBox.SelectionLength = lastSelectionLength
                End If
            End If
        End Sub

        Private Sub InputTextBox_TextInput(sender As Object, e As TextCompositionEventArgs)
            If allowNumbersOnly AndAlso Not IsNumeric(e.Text) Then
                Select Case e.Text
                    Case "-", ".", vbBack
                    Case Else
                        e.Handled = True
                        Beep()
                End Select
            End If
        End Sub

        Private Sub InputTextBox_SelectionChanged(sender As Object, e As RoutedEventArgs)
            If Not exitTextChange Then
                lastSelectionStart = InputTextBox.SelectionStart
                lastSelectionLength = InputTextBox.SelectionLength
            End If
        End Sub

        <ExMethod>
        Friend Sub AppendFormatted(
                           text As Primitive,
                           fontName As Primitive,
                           fontSize As Primitive,
                           isBold As Primitive,
                           isItalic As Primitive,
                           isUnderlined As Primitive,
                           foreColor As Primitive,
                           backColor As Primitive)

            Dim txtRun As New Run()
            txtRun.Text = text

            Dim _fontName = fontName.AsString()
            _fontName = If(_fontName = "", TextWindow.FontName.AsString, _fontName)
            txtRun.FontFamily = New Media.FontFamily(_fontName)

            txtRun.FontSize = CDbl(If(fontSize > 0, fontSize, TextWindow.FontSize))

            Dim _bold = CBool(If(isBold.AsString() = "", TextWindow.FontBold, isBold))
            txtRun.FontWeight = If(_bold, FontWeights.Bold, FontWeights.Normal)

            Dim _italic = CBool(If(isItalic.AsString() = "", TextWindow.FontItalic, isItalic))
            txtRun.FontStyle = If(_italic, FontStyles.Italic, FontStyles.Normal)

            Dim _underlined = CBool(If(isUnderlined.AsString() = "", TextWindow.FontUnderlined, isUnderlined))
            If _underlined Then txtRun.TextDecorations = TextDecorations.Underline

            Dim _foreColor = If(foreColor = Colors.None, TextWindow.ForeColor, foreColor)
            txtRun.Foreground = Color.GetBrush(_foreColor)

            Dim _backColor = If(backColor = Colors.None, TextWindow.BackColor, backColor)
            txtRun.Background = Color.GetBrush(_backColor)

            paragraph.Inlines.Add(txtRun)
        End Sub

        Friend Function ReadDate() As Primitive
            ShowDatePicker()
            Do Until InputDatePicker.Visibility = Visibility.Collapsed
                DoEvents()
                If isWindowClosed Then Return ""
            Loop

            Dim d = InputDatePicker.SelectedDate
            Return If(d.HasValue(),
                New Primitive(d.Value.Ticks, NumberType.Date),
                New Primitive("")
            )
        End Function

        Private Sub ShowDatePicker()
            If Me.WindowState = WindowState.Minimized Then
                SystemCommands.RestoreWindow(Me)
            End If
            Me.Activate()
            Dim rect = GetCaretPos()

            With InputDatePicker
                .Text = ""
                .FontFamily = New Media.FontFamily(TextWindow.FontName)
                .FontSize = TextWindow._FontSize
                .FontWeight = If(CBool(TextWindow.FontBold), FontWeights.Bold, FontWeights.Normal)
                If ConsoleBox.FlowDirection = FlowDirection.RightToLeft Then
                    .Margin = New Thickness(0, rect.Y, rect.X, 0)
                Else
                    .Margin = New Thickness(rect.X, rect.Y, 0, 0)
                End If

                .Visibility = Visibility.Visible
                .Dispatcher.BeginInvoke(DispatcherPriority.Render,
                    Sub()
                        .Focus()
                        Dim textBox = TryCast(.Template.FindName("PART_TextBox", InputDatePicker), DatePickerTextBox)
                        If textBox IsNot Nothing Then textBox.Focus()
                    End Sub)
            End With
        End Sub

        Private Sub InputDatePicker_PreviewKeyDown(sender As Object, e As KeyEventArgs)
            If e.Key = Key.Enter Then
                Dispatcher.BeginInvoke(
                      Sub() CommitDatePicker()
                )
            ElseIf e.Key = Key.Escape Then
                InputDatePicker.SelectedDate = Nothing
            End If
        End Sub

        Private Sub CommitDatePicker()
            InputDatePicker.Visibility = Visibility.Collapsed
            allowNumbersOnly = False
            Dim d = InputDatePicker.SelectedDate
            Dim d2 = If(d.HasValue(),
                    New Primitive(d.Value.Ticks, NumberType.Date),
                    New Primitive("")
            )
            WriteLine([Date].GetShortDate(d2))
            ConsoleBox.Focus()
            ConsoleBox.CaretPosition = ConsoleBox.Document.ContentEnd
        End Sub

        Private Sub ConsoleWindow_IsVisibleChanged(sender As Object, e As DependencyPropertyChangedEventArgs) Handles Me.IsVisibleChanged
            If Me.IsVisible Then
                _threadTimer.Change(1000, 1000)
            Else
                _threadTimer.Change(-1, -1)
            End If
        End Sub
    End Class
End Namespace
