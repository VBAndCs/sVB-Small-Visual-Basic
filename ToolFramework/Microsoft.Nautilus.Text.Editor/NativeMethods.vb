Imports System.Runtime.InteropServices
Imports System.Security

Namespace Microsoft.Nautilus.Text.Editor

    Friend NotInheritable Class NativeMethods
        Public Structure POINT
            Public x As Integer

            Public y As Integer
        End Structure

        Public Structure RECT_Renamed
            Public left As Integer

            Public top As Integer

            Public right As Integer

            Public bottom As Integer
        End Structure

        Public Structure MONITORINFO
            Public cbSize As Integer

            Public rcMonitor As RECT_Renamed

            Public rcWork As RECT_Renamed

            Public dwFlags As Integer
        End Structure

        Public Structure COMPOSITIONFORM
            Public dwStyle As Integer

            Public ptCurrentPos As POINT

            Public rcArea As RECT_Renamed
        End Structure

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
        Public Class LOGFONT
            Public lfHeight As Integer
            Public lfWidth As Integer
            Public lfEscapement As Integer
            Public lfOrientation As Integer
            Public lfWeight As Integer
            Public lfItalic As Byte
            Public lfUnderline As Byte
            Public lfStrikeOut As Byte
            Public lfCharSet As Byte
            Public lfOutPrecision As Byte
            Public lfClipPrecision As Byte
            Public lfQuality As Byte
            Public lfPitchAndFamily As Byte
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
            Public lfFaceName As String
        End Class

        <ComImport>
        <Guid("aa80e801-2021-11d2-93e0-0060b067b86e")>
        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
        Friend Interface ITfThreadMgr
            <SecurityCritical>
            <SuppressUnmanagedCodeSecurity>
            Sub Activate(<Out> ByRef clientId As Integer)
            <SecurityCritical>
            <SuppressUnmanagedCodeSecurity>
            Sub Deactivate()
            <SecurityCritical>
            <SuppressUnmanagedCodeSecurity>
            Sub CreateDocumentMgr(<Out> ByRef docMgr As Object)

            Sub EnumDocumentMgrs(<Out> ByRef enumDocMgrs As Object)

            Sub GetFocus(<Out> ByRef docMgr As IntPtr)
            <SecurityCritical>
            <SuppressUnmanagedCodeSecurity>
            Sub SetFocus(docMgr As IntPtr)

            Sub AssociateFocus(hwnd As IntPtr, newDocMgr As Object, <Out> ByRef prevDocMgr As Object)

            Sub IsThreadFocus(<Out> <MarshalAs(UnmanagedType.Bool)> ByRef isFocus As Boolean)
            <PreserveSig>
            <SuppressUnmanagedCodeSecurity>
            <SecurityCritical>
            Function GetFunctionProvider(ByRef classId As Guid, <Out> ByRef funcProvider As Object) As Integer

            Sub EnumFunctionProviders(<Out> ByRef enumProviders As Object)
            <SuppressUnmanagedCodeSecurity>
            <SecurityCritical>
            Sub GetGlobalCompartment(<Out> ByRef compartmentMgr As Object)
        End Interface

        Public Const MONITOR_DEFAULTTONEAREST As Integer = 2
        Public Const WM_IME_STARTCOMPOSITION As Integer = 269
        Public Const CFS_POINT As Integer = 2

        <DllImport("user32.dll")>
        Public Shared Function ClientToScreen(hWnd1 As IntPtr, ByRef point1 As POINT) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport("user32.dll")>
        Public Shared Function MonitorFromWindow(hwnd As IntPtr, dwFlags As Integer) As IntPtr
        End Function

        <DllImport("user32.dll")>
        Public Shared Function MonitorFromPoint(pt As POINT, dwFlags As Integer) As IntPtr
        End Function

        <DllImport("user32.dll")>
        Public Shared Function GetMonitorInfo(hMonitor As IntPtr, ByRef lpmi As MONITORINFO) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport("msctf.dll")>
        Friend Shared Function TF_CreateThreadMgr(<Out> ByRef threadMgr As ITfThreadMgr) As Integer
        End Function

        <DllImport("imm32.dll")>
        Friend Shared Function ImmGetContext(hWnd1 As IntPtr) As IntPtr
        End Function

        <DllImport("imm32.dll")>
        Friend Shared Function ImmSetCompositionWindow(hIMC As IntPtr, ptr As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport("imm32.dll")>
        Friend Shared Function ImmReleaseContext(hWnd1 As IntPtr, hIMC As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport("imm32.dll")>
        Friend Shared Function ImmSetCompositionFontW(hIMC As IntPtr, lplf As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

    End Class

End Namespace
