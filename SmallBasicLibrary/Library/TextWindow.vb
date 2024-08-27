Imports System.Text
Imports Microsoft.SmallVisualBasic.Library.Internal

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
                Return Console.ForegroundColor.ToString()
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Try
                    Value = WinForms.Color.GetName(Value)
                    Console.ForegroundColor = CType([Enum].Parse(GetType(ConsoleColor), Value, ignoreCase:=True), ConsoleColor)
                Catch

                End Try
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the background color of the text to be output in the text window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Property BackgroundColor As Primitive
            Get
                VerifyAccess()
                Return Console.BackgroundColor.ToString()
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Try
                    Value = WinForms.Color.GetName(Value)
                    Console.BackgroundColor = CType([Enum].Parse(GetType(ConsoleColor), Value, ignoreCase:=True), ConsoleColor)
                Catch

                End Try
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the cursor's column position on the text window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property CursorLeft As Primitive
            Get
                VerifyAccess()
                Return New Primitive(Console.CursorLeft)
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Console.CursorLeft = Value
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the cursor's row position on the text window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property CursorTop As Primitive
            Get
                VerifyAccess()
                Return New Primitive(Console.CursorTop)
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Console.CursorTop = Value
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the Left position of the Text Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Left As Primitive
            Get
                VerifyAccess()
                Dim consoleWindow As IntPtr = NativeHelper.GetConsoleWindow()
                Dim lpRect As RECT = Nothing
                NativeHelper.GetWindowRect(consoleWindow, lpRect)
                Return lpRect.Left
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Dim consoleWindow As IntPtr = NativeHelper.GetConsoleWindow()
                Dim lpRect As RECT = Nothing
                NativeHelper.GetWindowRect(consoleWindow, lpRect)
                NativeHelper.SetWindowPos(consoleWindow, IntPtr.Zero, Value, lpRect.Top, 0, 0, 1UL)
            End Set
        End Property

        <HideFromIntellisense>
        Public Shared Sub ClearIfLoaded()
            Dim consoleWindow = NativeHelper.GetConsoleWindow()
            If consoleWindow <> IntPtr.Zero Then Console.Clear()
        End Sub

        ''' <summary>
        ''' Gets or sets the Title for the text window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Property Title As Primitive
            Get
                VerifyAccess()
                Return New Primitive(Console.Title)
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Console.Title = Value
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the Top position of the Text Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Top As Primitive
            Get
                VerifyAccess()
                Dim consoleWindow As IntPtr = NativeHelper.GetConsoleWindow()
                Dim lpRect As RECT = Nothing
                NativeHelper.GetWindowRect(consoleWindow, lpRect)
                Return lpRect.Top
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Dim consoleWindow As IntPtr = NativeHelper.GetConsoleWindow()
                Dim lpRect As RECT = Nothing
                NativeHelper.GetWindowRect(consoleWindow, lpRect)
                NativeHelper.SetWindowPos(consoleWindow, IntPtr.Zero, lpRect.Left, Value, 0, 0, 1UL)
            End Set
        End Property

        <HideFromIntellisense>
        Public Shared Sub Close()
            Dim consoleWindow = NativeHelper.GetConsoleWindow()
            If consoleWindow <> IntPtr.Zero Then
                NativeHelper.ShowWindow(consoleWindow, 0)
                NativeHelper.FreeConsole()
            End If
        End Sub

        ''' <summary>
        ''' Shows the Text window to enable interactions with it.
        ''' </summary>
        Public Shared Sub Show()
            If Not _windowVisible Then
                Dim consoleWindow = NativeHelper.GetConsoleWindow()
                If consoleWindow = IntPtr.Zero Then
                    NativeHelper.AllocConsole()
                    consoleWindow = NativeHelper.GetConsoleWindow()
                    'If SmallBasicApplication.IsDebugging Then
                    Dim systemMenu = NativeHelper.GetSystemMenu(consoleWindow, False)
                    NativeHelper.DeleteMenu(systemMenu, NativeHelper.SC_CLOSE, NativeHelper.MF_BYCOMMAND)
                    'End If
                End If

                    NativeHelper.ShowWindow(consoleWindow, 5)
                _windowVisible = True
            End If
        End Sub


        ''' <summary>
        ''' Hides the text window.  Content is perserved when the window is shown again.
        ''' </summary>
        Public Shared Sub Hide()
            If _windowVisible Then
                Dim consoleWindow As IntPtr = NativeHelper.GetConsoleWindow()
                If consoleWindow <> IntPtr.Zero Then
                    SmallBasicApplication.BeginInvoke(Sub() NativeHelper.ShowWindow(consoleWindow, 0))
                End If
                _windowVisible = False
            End If
        End Sub

        ''' <summary>
        ''' Clears the TextWindow.
        ''' </summary>
        Public Shared Sub Clear()
            VerifyAccess()
            Console.Clear()
        End Sub

        ''' <summary>
        ''' Waits for user input before returning.
        ''' </summary>
        Public Shared Sub Pause()
            VerifyAccess()
            Console.WriteLine("Press any key to continue...")
            Try
                Console.ReadKey(intercept:=True)
            Catch
            End Try
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
            Try
                If _windowVisible Then
                    If addLine Then
                        Console.WriteLine(vbCrLf & "Press any key to close the window...")
                    Else
                        Console.WriteLine("Press any key to close the window...")
                    End If
                    Console.ReadKey(intercept:=True)
                End If

                If WinForms.Forms._forms.Count = 0 AndAlso Not GraphicsWindow._windowVisible Then
                    Program.End()
                Else
                    Hide()
                End If

            Catch
                Console.WriteLine("An error occured.")
                Console.WriteLine("If you are debugging this sVB projet in VS.NET in its debugging mode, thee console window will nor work correctly.")
                Console.WriteLine("Stop the VS debugger the Press Ctrl+F5 to run sVB im VS without debugging, then debug your project in sVB.")
            End Try
        End Sub

        ''' <summary>
        ''' Waits for user input before returning.
        ''' </summary>
        Public Shared Sub PauseWithoutMessage()
            VerifyAccess()
            Try
                Console.ReadKey(intercept:=True)
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Reads a line of text from the text window.  This function will not return until the user hits ENTER.
        ''' Use ? as a shortcut name for this method.
        ''' </summary>
        ''' <returns>The text that was read from the text window</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function Read() As Primitive
            VerifyAccess()
            Dim x = Console.ReadLine()
            Return New Primitive(x)
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
            Return New String(Console.ReadKey(intercept:=True).KeyChar, 1)
        End Function

        ''' <summary>
        ''' Reads a number from the text window.  This function will not return until the user hits ENTER.
        ''' Use #? as a shortcut name for this method.
        ''' </summary>
        ''' <returns>The number that was read from the text window</returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function ReadNumber() As Primitive
            VerifyAccess()
            Dim sbNumber As New StringBuilder
            Dim exists As Boolean = False
            Dim count As Integer = 0

            Do
                Dim keyInfo = Console.ReadKey(intercept:=True)
                Dim c As Char = keyInfo.KeyChar
                Dim isValid As Boolean = False
                If (c = "-"c AndAlso count = 0) OrElse
                        (c >= "0"c AndAlso c <= "9"c) Then
                    isValid = True
                ElseIf c = "."c AndAlso Not exists Then
                    exists = True
                    isValid = True
                End If

                If isValid Then
                    Console.Write(c)
                    sbNumber.Append(c)
                    count += 1
                ElseIf count > 0 AndAlso keyInfo.Key = ConsoleKey.Backspace Then
                    Console.CursorLeft -= 1
                    Console.Write(" ")
                    Console.CursorLeft -= 1
                    count -= 1
                    c = sbNumber(count)
                    If c = "."c Then exists = False
                    sbNumber.Remove(count, 1)
                ElseIf keyInfo.Key = ConsoleKey.Enter Then
                    Exit Do
                End If
            Loop

            Console.WriteLine()
            If sbNumber.Length = 0 Then
                Return New Primitive(0)
            Else
                Return New Primitive(sbNumber.ToString())
            End If
        End Function

        ''' <summary>
        ''' Writes text or number to the text window.  A new line character will be appended to the output, so that the next time something is written to the text window, it will go in a new line.
        ''' Use ? as a shortcut name for this method.
        ''' </summary>
        ''' <param name="data">The text or number to write to the text window.</param>
        Public Shared Sub WriteLine(data As Primitive)
            VerifyAccess()
            Console.WriteLine(data.AsString())
            addLine = False
        End Sub

        ''' <summary>
        ''' Writes the items of the given array to the TextWindow, and appends a new line after each of them.
        ''' Use ? as a shortcut name for this method, followed by an array initializer or a direct comman separated values
        ''' </summary>
        ''' <param name="lines">An array of text lines</param>
        Public Shared Sub WriteLines(lines As Primitive)
            VerifyAccess()
            If lines.IsArray Then
                Dim map = lines._arrayMap
                If map Is Nothing Then
                    Console.WriteLine()
                Else
                    For Each line In map.Values
                        Console.WriteLine(CStr(line))
                    Next
                End If
            Else
                Console.WriteLine(CStr(lines))
            End If
            addLine = False
        End Sub

        ''' <summary>
        ''' Writes text or number to the text window.  Unlike WriteLine, this will not append a new line character, which means, anything written to the text window after this call will be on the same line.
        ''' Use ?: as a shortcut name for this method.
        ''' </summary>
        ''' <param name="data">The text or number to write to the text window</param>
        Public Shared Sub Write(data As Primitive)
            VerifyAccess()
            Dim x As String = data
            Console.Write(x)
            addLine = True
        End Sub

        ''' <summary>
        ''' Formats the given text by substituting the placeholders by the given values, and writes the resulted text to the current line of the Text Window.
        ''' </summary>
        ''' <param name="text">The string to Format. Use [1], [2],... [n] in the string, to refer the values[1], values[2], ... values[n]</param>
        ''' <param name="values">An array its elements will be used to replace [1], [2],... [n] placeholders if found in the text</param>
        Public Shared Sub WriteFormated(text As Primitive, values As Primitive)
            Write(Library.Text.Format(text, values))
        End Sub

        ''' <summary>
        ''' Verifies if the access to text Window has been made yet
        ''' </summary>
        Private Shared Sub VerifyAccess()
            If Not _windowVisible Then Show()
        End Sub
    End Class
End Namespace
