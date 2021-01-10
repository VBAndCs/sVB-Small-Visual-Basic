Imports System
Imports System.Runtime.InteropServices

Namespace Microsoft.SmallBasic.Utility
    Public Module NativeMethods
        Public Const SC_CLOSE As Integer = 61536
        Public Const MF_ENABLED As Integer = 0
        Public Const MF_GREYED As Integer = 1

        <DllImport("user32")>
        Public Function GetSystemMenu(ByVal hWnd As IntPtr, ByVal revert As Integer) As IntPtr
        End Function

        <DllImport("user32")>
        Public Function EnableMenuItem(ByVal hWndMenu As IntPtr, ByVal itemID As Integer, ByVal enable As Integer) As Integer
        End Function
    End Module
End Namespace
