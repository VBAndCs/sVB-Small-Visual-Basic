Imports System.Windows.Interop
Imports System.Runtime.InteropServices
Imports System.IO

Friend Class CursorHelper
    Private Structure IconInfo
        Public fIcon As Boolean
        Public xHotspot As Integer
        Public yHotspot As Integer
        Public hbmMask As IntPtr
        Public hbmColor As IntPtr
    End Structure


    <DllImport("user32.dll")>
    Private Shared Function CreateIconIndirect(ByRef icon As IconInfo) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetIconInfo(hIcon As IntPtr, ByRef pIconInfo As IconInfo) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Private Shared Function InternalCreateCursor(
                 bmp As System.Drawing.Bitmap,
                 xHotSpot As Integer,
                 yHotSpot As Integer
            ) As Cursor

        Dim iconInfo As New IconInfo()
        GetIconInfo(bmp.GetHicon(), iconInfo)
        iconInfo.xHotspot = xHotSpot
        iconInfo.yHotspot = yHotSpot
        iconInfo.fIcon = False

        Try
            Dim ptr = CreateIconIndirect(iconInfo)
            Dim handle As New SafeIconHandle(ptr)
            Return CursorInteropHelper.Create(handle)
        Catch ex As Exception

        End Try

        Return Nothing
    End Function

    Public Shared Function CreateCursor(
                       element As UIElement,
                       xHotSpot As Integer,
                       yHotSpot As Integer
                ) As Cursor

        Try
            Dim Brdr As New Border
            Brdr.Child = element
            Brdr.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
            Brdr.Arrange(New Rect(0, 0, element.DesiredSize.Width, element.DesiredSize.Height))

            Dim Width As Integer = Brdr.ActualWidth
            Dim Height As Integer = Brdr.ActualHeight

            Dim rtb = New RenderTargetBitmap(Width, Height, 96, 96, PixelFormats.Pbgra32)
            rtb.Render(Brdr)

            Dim encoder As New PngBitmapEncoder()
            encoder.Frames.Add(BitmapFrame.Create(rtb))

            Dim ms As New MemoryStream()
            encoder.Save(ms)

            Dim bmp As New System.Drawing.Bitmap(ms)

            ms.Close()
            ms.Dispose()
            rtb = Nothing

            If xHotSpot = -1 Then xHotSpot = Width / 2
            If yHotSpot = -1 Then yHotSpot = Height / 2

            Dim cur As Cursor = InternalCreateCursor(bmp, xHotSpot, yHotSpot)

            bmp.Dispose()
            Brdr.Child = Nothing
            Return cur

        Catch
        End Try
        Return Nothing

    End Function
End Class

