Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Documents
Imports System.Threading.Tasks
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Library.Internal
Imports System.Windows.Controls

Namespace WinForms
    Partial Public Class ConsoleWindow
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
            Me.Activate()
        End Sub

        Public Sub WriteLine(text As String)
            If text <> "" Then Write(text)
            paragraph.Inlines.Add(New LineBreak())
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
            Dim caretPosition = ConsoleBox.CaretPosition
            Dim textPointer = caretPosition.GetCharacterRect(LogicalDirection.Forward)

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
                    .Margin = New Thickness(0, textPointer.Y, textPointer.X, 0)
                Else
                    .Margin = New Thickness(textPointer.X, textPointer.Y, 0, 0)
                End If
                .Visibility = Visibility.Visible
                .Focus()
            End With
        End Sub

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
        End Sub

        Friend Sub WaitForAnyKey()
            isKeyPressed = False
            Do
                DoEvents()
                If isKeyPressed OrElse isWindowClosed Then Return
            Loop
        End Sub

        Private Sub Window_PreviewKeyDown(sender As Object, e As KeyEventArgs)
            isKeyPressed = True
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
            Dim caretPosition = ConsoleBox.CaretPosition
            Dim textPointer = caretPosition.GetCharacterRect(LogicalDirection.Forward)

            With InputDatePicker
                .Text = ""
                .FontFamily = New Media.FontFamily(TextWindow.FontName)
                .FontSize = TextWindow._FontSize
                .FontWeight = If(CBool(TextWindow.FontBold), FontWeights.Bold, FontWeights.Normal)
                If ConsoleBox.FlowDirection = FlowDirection.RightToLeft Then
                    .Margin = New Thickness(0, textPointer.Y, textPointer.X, 0)
                Else
                    .Margin = New Thickness(textPointer.X, textPointer.Y, 0, 0)
                End If
                .Visibility = Visibility.Visible
                .Focus()
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
    End Class
End Namespace
