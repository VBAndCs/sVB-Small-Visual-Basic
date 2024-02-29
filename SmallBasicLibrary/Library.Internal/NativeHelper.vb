Imports System.Runtime.InteropServices
Imports System.Text

Namespace Library.Internal
    ''' <summary>
    ''' A private static helper for calling Native APIs
    ''' </summary>
    Friend NotInheritable Class NativeHelper
        Public Const STD_INPUT_HANDLE As Integer = -10
        Public Const STD_OUTPUT_HANDLE As Integer = -11
        Public Const STD_ERROR_HANDLE As Integer = -12
        Public Const SPI_SETDESKWALLPAPER As Integer = 20
        Public Const SPIF_UPDATEINIFILE As Integer = 1
        Public Const SPIF_SENDWININICHANGE As Integer = 2
        Public Const GWL_STYLE As Integer = -16
        Public Const WS_MINIMIZEBOX As UInteger = 131072UL
        Public Const WS_MAXIMIZEBOX As UInteger = 65536UL
        Public Const WS_CAPTION As UInteger = 12582912UL
        Public Const SWP_NOSIZE As Integer = 1
        Public Const MF_BYCOMMAND As Integer = &H0
        Public Const SC_CLOSE As Integer = &HF060

        <DllImport("user32.dll")>
        Public Shared Function DeleteMenu(ByVal hMenu As IntPtr, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
        End Function

        <DllImport("user32.dll")>
        Public Shared Function GetSystemMenu(ByVal hWnd As IntPtr, ByVal bRevert As Boolean) As IntPtr
        End Function

        <DllImport("kernel32")>
        Public Shared Function AllocConsole() As Boolean
        End Function

        <DllImport("kernel32.dll")>
        Public Shared Function FreeConsole() As Boolean
        End Function

        <DllImport("user32.dll")>
        Public Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
        End Function

        <DllImport("kernel32.dll")>
        Public Shared Function GetConsoleWindow() As IntPtr
        End Function

        <DllImport("kernel32.dll")>
        Public Shared Function GetConsoleTitle(<Out> lpConsoleTitle As StringBuilder, nSize As UInteger) As UInteger
        End Function

        <DllImport("kernel32.dll")>
        Public Shared Function SetConsoleTitle(lpConsoleTitle As String) As Boolean
        End Function

        <DllImport("kernel32.dll")>
        Public Shared Function SetConsoleTextAttribute(hConsoleOutput As IntPtr, wAttributes As UShort) As Boolean
        End Function

        <DllImport("kernel32.dll")>
        Public Shared Function GetStdHandle(nStdHandle As Integer) As IntPtr
        End Function

        <DllImport("user32.dll")>
        Public Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Public Shared Function SystemParametersInfo(uAction As Integer, uParam As Integer, lpvParam As String, fuWinIni As Integer) As Integer
        End Function

        <DllImport("user32.dll")>
        Public Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As UInteger
        End Function

        <DllImport("user32.dll")>
        Public Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As UInteger) As Integer
        End Function

        <DllImport("user32.dll")>
        Public Shared Function AdjustWindowRect(ByRef lpRect As RECT, dwStyle As UInteger, bMenu As Boolean) As Boolean
        End Function

        <DllImport("user32.dll")>
        Public Shared Function GetWindowRect(hWnd As IntPtr, <Out> ByRef lpRect As RECT) As Boolean
        End Function

        <DllImport("user32.dll")>
        Public Shared Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr, X As Integer, Y As Integer, cx As Integer, cy As Integer, uFlags As UInteger) As Boolean
        End Function

        <DllImport("winmm.dll")>
        Public Shared Function midiOutOpen(ByRef midiOut As Integer, uDeviceID As Integer, dwCallback As Integer, dwCallbackInstance As Integer, dwFlags As UInteger) As Integer
        End Function

        <DllImport("winmm.dll")>
        Public Shared Function midiOutShortMsg(midiOut As Integer, dwMsg As UInteger) As Integer
        End Function
    End Class
End Namespace
