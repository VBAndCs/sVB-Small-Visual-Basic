Imports System
Imports System.Windows.Interop
Imports Microsoft.Win32.SafeHandles
Imports System.Runtime.InteropServices
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.IO
Imports System.Windows

Friend Class CursorHelper
    Private Structure IconInfo
        Public fIcon As Boolean
        Public xHotspot As Integer
        Public yHotspot As Integer
        Public hbmMask As IntPtr
        Public hbmColor As IntPtr
    End Structure


    <DllImport("user32.dll")> _
    Private Shared Function CreateIconIndirect(ByRef icon As IconInfo) As IntPtr
    End Function

    <DllImport("user32.dll")> _
    Private Shared Function GetIconInfo(ByVal hIcon As IntPtr, ByRef pIconInfo As IconInfo) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function


    Private Shared Function InternalCreateCursor(ByVal bmp As System.Drawing.Bitmap, ByVal xHotSpot As Integer, ByVal yHotSpot As Integer) As Cursor
        Dim tmp As New IconInfo()
        GetIconInfo(bmp.GetHicon(), tmp)
        tmp.xHotspot = xHotSpot
        tmp.yHotspot = yHotSpot
        tmp.fIcon = False

        Dim ptr As IntPtr = CreateIconIndirect(tmp)
        Dim handle As New SafeFileHandle(ptr, True)
        Return CursorInteropHelper.Create(handle)
    End Function

    Public Shared Function CreateCursor(ByVal element As UIElement, ByVal xHotSpot As Integer, ByVal yHotSpot As Integer) As Cursor
        Dim Brdr As New Border
        Brdr.Child = element
        Brdr.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
        Brdr.Arrange(New Rect(0, 0, element.DesiredSize.Width, element.DesiredSize.Height))

        Dim Width = Brdr.ActualWidth
        Dim Height = Brdr.ActualHeight

        Dim rtb = New RenderTargetBitmap(CInt(Width), CInt(Height), 96, 96, PixelFormats.Pbgra32)
        rtb.Render(Brdr)

        Dim encoder As New PngBitmapEncoder()
        encoder.Frames.Add(BitmapFrame.Create(rtb))

        Dim ms As New MemoryStream()
        encoder.Save(ms)

        Dim bmp As New System.Drawing.Bitmap(ms)

        ms.Close()
        ms.Dispose()

        If xHotSpot = -1 Then xHotSpot = Width / 2
        If yHotSpot = -1 Then yHotSpot = Height / 2

        Dim cur As Cursor = InternalCreateCursor(bmp, xHotSpot, yHotSpot)

        bmp.Dispose()
        Brdr.Child = Nothing
        Return cur
    End Function
End Class
