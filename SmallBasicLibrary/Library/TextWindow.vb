Imports System.Text
Imports Microsoft.SmallVisualBasic.Library.Internal
Imports Microsoft.SmallVisualBasic.WinForms

Namespace Library
    ''' <summary>
    ''' Provides text-related input and output functionalities.  For example using this class, it is possible to write or read some text or number to and from the text-based text window.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class TextWindow
        Friend Shared _windowVisible As Boolean
        Private Shared addLine As Boolean = False


        ''' <summary>
        ''' Gets or sets the foreground color of the text to be output in the text window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Property ForegroundColor As Primitive
            Get
                VerifyAccess()
                SmallBasicApplication.Invoke(
                      Sub() ForegroundColor = New Primitive(WinForms.Color.GetHexaName(consoleWindow.ConsoleBox.Foreground))
                )
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                SmallBasicApplication.Invoke(
                      Sub() consoleWindow.ConsoleBox.Foreground = WinForms.Color.GetBrush(Value)
                )
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the background color of the text to be output in the text window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Property BackgroundColor As Primitive
            Get
                VerifyAccess()
                SmallBasicApplication.Invoke(
                      Sub() BackgroundColor = New Primitive(WinForms.Color.GetHexaName(consoleWindow.ConsoleBox.Background))
                )
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                SmallBasicApplication.Invoke(
                      Sub() consoleWindow.ConsoleBox.Background = WinForms.Color.GetBrush(Value)
                )
            End Set
        End Property

        Private Shared _rightToLeft As New Primitive(False)

        ''' <summary>
        ''' Gets or sets whether or not the console direction is from right to left.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property RightToLeft As Primitive
            Get
                Return _rightToLeft
            End Get

            Set(value As Primitive)
                _rightToLeft = value
                If _windowVisible Then
                    SmallBasicApplication.Invoke(
                      Sub() UpdateFlowDirection()
                )
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the cursor's column position on the text window.
        ''' </summary>
        <HideFromIntellisense>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property CursorLeft As Primitive
            Get
                VerifyAccess()
                Return 0
            End Get

            Set(Value As Primitive)
                VerifyAccess()
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the cursor's row position on the text window.
        ''' </summary>
        <HideFromIntellisense>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property CursorTop As Primitive
            Get
                VerifyAccess()
                Return 0
            End Get

            Set(Value As Primitive)
                VerifyAccess()
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the Left position of the Text Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Left As Primitive
            Get
                VerifyAccess()
                SmallBasicApplication.Invoke(
                      Sub() Left = New Primitive(consoleWindow.Left)
                )
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                SmallBasicApplication.Invoke(
                    Sub() consoleWindow.Left = Value
                )
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the Title for the text window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Property Title As Primitive
            Get
                VerifyAccess()
                SmallBasicApplication.Invoke(
                      Sub() Title = New Primitive(consoleWindow.Title)
                )
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                SmallBasicApplication.Invoke(
                    Sub() consoleWindow.Title = Value
                )
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the Top position of the Text Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Top As Primitive
            Get
                VerifyAccess()
                SmallBasicApplication.Invoke(
                      Sub() Top = New Primitive(consoleWindow.Top)
                )
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                SmallBasicApplication.Invoke(
                    Sub() consoleWindow.Top = Value
                )
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the font name that will be used to write to the TextWindow.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Property FontName As New Primitive("Consolas")

        ''' <summary>
        ''' Gets or sets the back color of the text that will be written to the TextWindow.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Property BackColor As Primitive = WinForms.Colors.None

        ''' <summary>
        ''' Gets or sets the font color of the text that will be written to the TextWindow.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Property ForeColor As Primitive = WinForms.Colors.None

        Friend Shared _FontSize As New Primitive(14 * 4 / 3)

        ''' <summary>
        ''' Gets or sets the font size that will be used to write to the TextWindow.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Property FontSize As Primitive
            Get
                Return _FontSize * 0.75
            End Get
            Set
                _FontSize = System.Math.Max(1, Value.AsDecimal) / 0.75
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets whether or not the text will be written in italic font.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property FontItalic As New Primitive(False)

        ''' <summary>
        ''' Gets or sets whether or not the text will be underlined.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property FontUnderlined As New Primitive(False)

        ''' <summary>
        ''' Gets or sets whether or not the text will be written in bold font.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property FontBold As New Primitive(False)

        <HideFromIntellisense>
        Public Shared Sub Close()
            SmallBasicApplication.Invoke(
                Sub()
                    If consoleWindow IsNot Nothing AndAlso Not consoleWindow.isWindowClosed Then
                        consoleWindow.Close()
                    End If
                End Sub)
        End Sub

        Friend Shared consoleWindow As WinForms.ConsoleWindow

        ''' <summary>
        ''' Shows the Text window to enable interactions with it.
        ''' </summary>
        Public Shared Sub Show()
            If Not _windowVisible Then
                SmallBasicApplication.Invoke(
                       Sub()
                           If consoleWindow Is Nothing OrElse consoleWindow.isWindowClosed Then
                               consoleWindow = New WinForms.ConsoleWindow
                           End If
                           consoleWindow.Show()
                           _windowVisible = True

                           UpdateFlowDirection()
                       End Sub)
            End If
        End Sub

        Private Shared Sub UpdateFlowDirection()
            Dim dir = If(CBool(_rightToLeft),
                                  System.Windows.FlowDirection.RightToLeft,
                                  System.Windows.FlowDirection.LeftToRight)

            Dim align = If(CBool(_rightToLeft),
                                  System.Windows.HorizontalAlignment.Right,
                                  System.Windows.HorizontalAlignment.Left)

            consoleWindow.ConsoleBox.FlowDirection = dir
            consoleWindow.InputTextBox.HorizontalAlignment = align
            consoleWindow.InputTextBox.FlowDirection = dir
            consoleWindow.InputDatePicker.HorizontalAlignment = align
            consoleWindow.InputDatePicker.FlowDirection = dir
        End Sub

        ''' <summary>
        ''' Hides the text window. Content is perserved when the window is shown again.
        ''' </summary>
        Public Shared Sub Hide()
            If _windowVisible Then
                SmallBasicApplication.Invoke(
                        Sub()
                            If consoleWindow IsNot Nothing AndAlso Not consoleWindow.isWindowClosed Then
                                consoleWindow.Hide()
                            End If
                        End Sub)
                _windowVisible = False
            End If
        End Sub

        ''' <summary>
        ''' Clears the TextWindow.
        ''' </summary>
        Public Shared Sub Clear()
            VerifyAccess()
            DoClear()
        End Sub

        ' It is a boolean function just to keep it out of the sVB lib
        ' but its return values has no use
        Public Shared Function DoClear() As Boolean
            SmallBasicApplication.Invoke(
                    Sub()
                        If consoleWindow IsNot Nothing AndAlso Not consoleWindow.isWindowClosed Then
                            consoleWindow.Clear()
                        End If
                    End Sub)
            Return False
        End Function

        ''' <summary>
        ''' Waits for user input before returning.
        ''' </summary>
        Public Shared Sub Pause()
            VerifyAccess()
            SmallBasicApplication.Invoke(
                Sub()
                    consoleWindow.WriteLine("Press any key to continue...")
                    consoleWindow.WaitForAnyKey()
                End Sub)
        End Sub

        ''' <summary>
        ''' Waits for user input only when the TextWindow is already open.
        ''' </summary>
        Public Shared Sub PauseIfVisible()
            If _windowVisible Then Pause()
        End Sub

        ''' <summary>
        ''' Waits for user to press any key to close the window, otherwise it closes it directly.
        ''' </summary>
        Public Shared Sub PauseThenClose()
            If _windowVisible Then
                SmallBasicApplication.Invoke(
                    Sub()
                        consoleWindow.WriteLine(If(addLine, vbCrLf, "") & "Press any key to close the window...")
                        consoleWindow.WaitForAnyKey()
                        If WinForms.Forms._forms.Count = 0 AndAlso Not GraphicsWindow._windowVisible Then
                            Program.End()
                        Else
                            Close()
                        End If
                    End Sub)
            End If
        End Sub

        ''' <summary>
        ''' Waits for user input before returning.
        ''' </summary>
        Public Shared Sub PauseWithoutMessage()
            VerifyAccess()
            SmallBasicApplication.Invoke(
                  Sub() consoleWindow.WaitForAnyKey()
            )
        End Sub

        ''' <summary>
        ''' Reads a line of text from the text window.  This function will not return until the user hits ENTER.
        ''' Use ? as a shortcut name for this method.
        ''' </summary>
        ''' <returns>The text that was read from the text window</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function Read() As Primitive
            VerifyAccess()
            SmallBasicApplication.Invoke(
                   Sub() Read = consoleWindow.ReadLine()
            )
        End Function

        ''' <summary>
        ''' Reads a single character from the text window.  
        ''' </summary>
        ''' <returns>
        ''' The character that was read from the text window.
        ''' </returns>
        <HideFromIntellisense>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function ReadKey() As Primitive
            VerifyAccess()
            SmallBasicApplication.Invoke(
                   Sub() ReadKey = consoleWindow.ReadKey()
            )
        End Function

        ''' <summary>
        ''' Reads a number from the text window.  This function will not return until the user hits ENTER.
        ''' Use #? as a shortcut name for this method.
        ''' </summary>
        ''' <returns>The number that was read from the text window</returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function ReadNumber() As Primitive
            VerifyAccess()
            SmallBasicApplication.Invoke(
                   Sub() ReadNumber = consoleWindow.ReadNumber()
            )
        End Function

        ''' <summary>
        ''' Reads a number from the text window.  This function will not return until the user hits ENTER.
        ''' Use #? as a shortcut name for this method.
        ''' </summary>
        ''' <returns>The number that was read from the text window</returns>
        <WinForms.ReturnValueType(VariableType.Date)>
        Public Shared Function ReadDate() As Primitive
            VerifyAccess()
            SmallBasicApplication.Invoke(
                   Sub() ReadDate = consoleWindow.ReadDate()
            )
        End Function


        ''' <summary>
        ''' Writes text or number to the text window.  A new line character will be appended to the output, so that the next time something is written to the text window, it will go in a new line.
        ''' Use ? as a shortcut name for this method.
        ''' </summary>
        ''' <param name="data">The text or number to write to the text window.</param>
        Public Shared Sub WriteLine(data As Primitive)
            VerifyAccess()
            SmallBasicApplication.Invoke(
                   Sub() consoleWindow.WriteLine(data.AsString())
            )
            addLine = False
        End Sub

        ''' <summary>
        ''' Writes the items of the given array to the TextWindow, and appends a new line after each of them.
        ''' Use ? as a shortcut name for this method, followed by an array initializer or a direct comman separated values
        ''' </summary>
        ''' <param name="lines">An array of text lines</param>
        Public Shared Sub WriteLines(lines As Primitive)
            VerifyAccess()
            SmallBasicApplication.Invoke(
                   Sub()
                       If lines.IsArray Then
                           Dim map = lines.ArrayMap
                           If map Is Nothing Then
                               consoleWindow.WriteLine("")
                           Else
                               For Each line In map.Values
                                   consoleWindow.WriteLine(CStr(line))
                               Next
                           End If
                       Else
                           consoleWindow.WriteLine(CStr(lines))
                       End If
                       addLine = False
                   End Sub)
        End Sub

        ''' <summary>
        ''' Writes text or number to the text window.  Unlike WriteLine, this will not append a new line character, which means, anything written to the text window after this call will be on the same line.
        ''' Use ?: as a shortcut name for this method.
        ''' </summary>
        ''' <param name="data">The text or number to write to the text window</param>
        Public Shared Sub Write(data As Primitive)
            VerifyAccess()
            SmallBasicApplication.Invoke(
                   Sub() consoleWindow.Write(data.AsString())
            )
            addLine = True
        End Sub

        ''' <summary>
        ''' Formats the given text by substituting the placeholders by the given values, and writes the resulted text to the current line of the Text Window.
        ''' </summary>
        ''' <param name="text">The string to format. Use [1], [2],... [n] in the string, to refer values[1], values[2], ... values[n] in the values array</param>
        ''' <param name="values">An array. Its elements will be used to replace [1], [2],... [n] placeholders if found in the text</param>
        Public Shared Sub WriteFormatted(text As Primitive, values As Primitive)
            Write(Library.Text.Format(text, values))
        End Sub

        ''' <summary>
        ''' Verifies if the access to text Window has been made yet
        ''' </summary>
        Private Shared Sub VerifyAccess()
            If Not _windowVisible Then Show()
        End Sub

        ''' <summary>
        ''' Writes the given text to the Text Window with the given formats.
        ''' </summary>
        ''' <param name="text">The text to write to the TW</param>
        ''' <param name="fontName">The name of the font to apply on the text. Send an empty string to use the TW.FontName</param>
        ''' <param name="fontSize">The font size of the text. Send 0 to use the TW.FontSize</param>
        ''' <param name="isBold">True to use a bold font, False otherwise. Send an empty string to use the TW.FontBold.</param>
        ''' <param name="isItalic">True to use an italic font, False otherwise. Send an empty string to use the TW.FontItalic.</param>
        ''' <param name="isUnderlined">True to draw aline under the text, False otherwise. Send an empty string to use TW.FontUnderlined.</param>
        ''' <param name="foreColor">The color of the text. Send Colors.None to use theTW.ForeColor.</param>
        ''' <param name="backColor">The background color of the text. Send Colors.None to use the TW.BackColor.</param>
        Public Shared Sub AppendFormatted(
                           text As Primitive,
                           fontName As Primitive,
                           fontSize As Primitive,
                           isBold As Primitive,
                           isItalic As Primitive,
                           isUnderlined As Primitive,
                           foreColor As Primitive,
                           backColor As Primitive)

            VerifyAccess()
            SmallBasicApplication.Invoke(
                  Sub() consoleWindow.AppendFormatted(text, fontName, fontSize, isBold, isItalic, isUnderlined, foreColor, backColor)
            )
            addLine = True
        End Sub


        ''' <summary>
        ''' Returns True if the Text Window is closed, or False otherwise.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared ReadOnly Property IsClosed As Primitive
            Get
                SmallBasicApplication.Invoke(
                      Sub() IsClosed = New Primitive(consoleWindow.isWindowClosed)
                )
            End Get
        End Property
    End Class
End Namespace
