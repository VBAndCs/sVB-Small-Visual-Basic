Imports System.Runtime.InteropServices
Imports Microsoft.Win32.SafeHandles

Class SafeIconHandle
    Inherits SafeHandleZeroOrMinusOneIsInvalid

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
    End Function

    Private Sub New()
        MyBase.New(True)
    End Sub

    Public Sub New(hIcon As IntPtr)
        MyBase.New(True)
        Me.SetHandle(hIcon)
    End Sub

    Protected Overrides Function ReleaseHandle() As Boolean
        Return DestroyIcon(Me.handle)
    End Function
End Class