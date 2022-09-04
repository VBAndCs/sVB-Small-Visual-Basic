Imports System.Text
Imports Microsoft.SmallBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' Provides text-related input and output functionalities.  For example using this class, it is possible to write or read some text or number to and from the text-based text window.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class TextWindow
        Friend Shared _windowVisible As Boolean

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

        ''' <summary>
        ''' Shows the Text window to enable interactions with it.
        ''' </summary>
        Public Shared Sub Show()
            If Not _windowVisible Then
                Dim consoleWindow As IntPtr = NativeHelper.GetConsoleWindow()
                If consoleWindow = IntPtr.Zero Then
                    NativeHelper.AllocConsole()
                End If
                NativeHelper.ShowWindow(consoleWindow, 8)
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
                    NativeHelper.ShowWindow(consoleWindow, 0)
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
            Console.Read()
            If WinForms.Forms._forms.Count = 0 AndAlso Not GraphicsWindow._windowVisible Then
                SmallBasicApplication.End()
            Else
                TextWindow.Hide()
            End If
        End Sub

        ''' <summary>
        ''' Waits for user input only when the TextWindow is already open.
        ''' </summary>
        Public Shared Sub PauseIfVisible()
            If _windowVisible Then
                Pause()
            End If
        End Sub

        ''' <summary>
        ''' Waits for user input before returning.
        ''' </summary>
        Public Shared Sub PauseWithoutMessage()
            VerifyAccess()
            Console.ReadKey(intercept:=True)
        End Sub

        ''' <summary>
        ''' Reads a line of text from the text window.  This function will not return until the user hits ENTER.
        ''' </summary>
        ''' <returns>
        ''' The text that was read from the text window
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function Read() As Primitive
            VerifyAccess()
            Return New Primitive(Console.ReadLine())
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
        ''' </summary>
        ''' <returns>
        ''' The number that was read from the text window
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function ReadNumber() As Primitive
            VerifyAccess()
            Dim stringBuilder1 As New StringBuilder
            Dim flag As Boolean = False
            Dim num As Integer = 0
            While True
                Dim consoleKeyInfo1 As ConsoleKeyInfo = Console.ReadKey(intercept:=True)
                Dim keyChar1 As Char = consoleKeyInfo1.KeyChar
                Dim flag2 As Boolean = False
                If keyChar1 = "-"c AndAlso num = 0 Then
                    flag2 = True
                ElseIf keyChar1 = "."c AndAlso Not flag Then
                    flag = True
                    flag2 = True
                ElseIf keyChar1 >= "0"c AndAlso keyChar1 <= "9"c Then
                    flag2 = True
                End If
                If flag2 Then
                    Console.Write(keyChar1)
                    stringBuilder1.Append(keyChar1)
                    num += 1
                ElseIf num > 0 AndAlso consoleKeyInfo1.Key = ConsoleKey.Backspace Then
                    Console.CursorLeft -= 1
                    Console.Write(" ")
                    Console.CursorLeft -= 1
                    num -= 1
                    keyChar1 = stringBuilder1(num)
                    If keyChar1 = "."c Then
                        flag = False
                    End If
                    stringBuilder1.Remove(num, 1)
                ElseIf consoleKeyInfo1.Key = ConsoleKey.Enter Then
                    Exit While
                End If
            End While
            Console.WriteLine()
            If stringBuilder1.Length = 0 Then
                Return New Primitive(0)
            End If

            Return New Primitive(stringBuilder1.ToString())
        End Function

        ''' <summary>
        ''' Writes text or number to the text window.  A new line character will be appended to the output, so that the next time something is written to the text window, it will go in a new line.
        ''' </summary>
        ''' <param name="data">
        ''' The text or number to write to the text window.
        ''' </param>
        Public Shared Sub WriteLine(data As Primitive)
            VerifyAccess()
            Console.WriteLine(CStr(data))
        End Sub

        ''' <summary>
        ''' Writes text or number to the text window.  Unlike WriteLine, this will not append a new line character, which means, anything written to the text window after this call will be on the same line.
        ''' </summary>
        ''' <param name="data">
        ''' The text or number to write to the text window
        ''' </param>
        Public Shared Sub Write(data As Primitive)
            VerifyAccess()
            Dim x As String = data
            Console.Write(x)
        End Sub

        ''' <summary>
        ''' Verifies if the access to text Window has been made yet
        ''' </summary>
        Private Shared Sub VerifyAccess()
            If Not _windowVisible Then
                Show()
            End If
        End Sub
    End Class
End Namespace
