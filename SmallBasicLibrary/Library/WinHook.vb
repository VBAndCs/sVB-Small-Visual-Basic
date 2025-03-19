Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Library
    Public Class WinHook
        Private Shared Started As Boolean = False
        Public Shared Sub Start()
            If Started Then Return

            Started = True
            Dim objCurrentModule = Process.GetCurrentProcess().MainModule
            objKeyboardProcess = New LowLevelKeyboardProc(AddressOf captureKey)
            ptrHook = SetWindowsHookEx(13, objKeyboardProcess, GetModuleHandle(objCurrentModule.ModuleName), 0)
        End Sub

        Public Shared Sub [Stop]()
            If Started Then
                Started = False
                UnhookWindowsHookEx(ptrHook)
            End If
        End Sub

        ' Structure contain information about low-level keyboard input event 
        <StructLayout(LayoutKind.Sequential)>
        Private Structure KBDLLHOOKSTRUCT
            Public key As Keys
            Public scanCode As Integer
            Public flags As Integer
            Public time As Integer
            Public extra As IntPtr
        End Structure

        'System level functions to be used for hook and unhook keyboard input  

        Private Delegate Function LowLevelKeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function SetWindowsHookEx(ByVal id As Integer, ByVal callback As LowLevelKeyboardProc, ByVal hMod As IntPtr, ByVal dwThreadId As UInteger) As IntPtr
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function UnhookWindowsHookEx(ByVal hook As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function CallNextHookEx(ByVal hook As IntPtr, ByVal nCode As Integer, ByVal wp As IntPtr, ByVal lp As IntPtr) As IntPtr
        End Function

        <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function GetModuleHandle(ByVal name As String) As IntPtr
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function GetAsyncKeyState(ByVal key As Keys) As Short
        End Function

        'Declaring Global objects     
        Private Shared ptrHook As IntPtr
        Private Shared objKeyboardProcess As LowLevelKeyboardProc

        Private Shared Function captureKey(nCode As Integer, wp As IntPtr, lp As IntPtr) As IntPtr
            If nCode >= 0 Then
                Dim objKeyInfo = CType(Marshal.PtrToStructure(lp, GetType(KBDLLHOOKSTRUCT)), KBDLLHOOKSTRUCT)
                If objKeyInfo.key = Keys.Escape AndAlso IsConsoleActive() Then
                    If WinForms.Forms._forms.Count = 0 AndAlso Not GraphicsWindow._windowVisible Then
                        Program.End()
                    Else
                        TextWindow.Hide()
                    End If
                End If
            End If

            Return CallNextHookEx(ptrHook, nCode, wp, lp)
        End Function

        Private Function HasAltModifier(ByVal flags As Integer) As Boolean
            Return (flags And &H20) = &H20
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetForegroundWindow() As IntPtr
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As UInteger) As UInteger
        End Function

        Shared Function IsConsoleActive() As Boolean
            Dim activeWindow = GetForegroundWindow()
            Dim consoleWindow = NativeHelper.GetConsoleWindow()
            If activeWindow <> consoleWindow Then Return False

            Dim processId As UInteger
            GetWindowThreadProcessId(activeWindow, processId)
            Return processId = CUInt(System.Diagnostics.Process.GetCurrentProcess().Id)
        End Function
    End Class
End Namespace