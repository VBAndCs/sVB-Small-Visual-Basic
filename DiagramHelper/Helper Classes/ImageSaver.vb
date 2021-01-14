Class ImageSaver
    Public DpiX As Integer = 300
    Public DpiY As Integer = 300
    Dim PixelFormat As PixelFormat = PixelFormats.Pbgra32

    Sub New()

    End Sub

    Sub New(DpiX As Integer, DpiY As Integer, PixelFormat As PixelFormat)
        Me.DpiX = DpiX
        Me.DpiY = DpiY
        Me.PixelFormat = PixelFormat
    End Sub

    Public Sub Save(ByVal Visual As FrameworkElement, ByVal fileName As String, Optional ShowSaveDialoge As Boolean = True)
        If ShowSaveDialoge Then
            Dim dlg As New Microsoft.Win32.SaveFileDialog()
            dlg.Title = "Save To Image"
            dlg.DefaultExt = ".JPEG" ' Default file extension
            dlg.FileName = IO.Path.GetFileNameWithoutExtension(fileName)
            dlg.Filter = "BMP Images|*.bmp|" &
                              "JPEG Images|*.jpg|" &
                              "GIF Images|*.gif|" &
                              "PNG Images|*.png|" &
                              "TIFF Images|*.tiff|" &
                              "WDP Images|*.wdp"

            dlg.FilterIndex = 2
            Dim result? As Boolean = dlg.ShowDialog()
            If result = True Then
                fileName = dlg.FileName
            Else
                Return
            End If
        End If

        Dim encoder As BitmapEncoder
        Select Case IO.Path.GetExtension(fileName).ToLower
            Case ".bmp"
                encoder = New BmpBitmapEncoder()
            Case ".png"
                encoder = New PngBitmapEncoder()
            Case ".gif"
                encoder = New GifBitmapEncoder()
            Case ".jpg"
                encoder = New JpegBitmapEncoder()
            Case ".tiff"
                encoder = New TiffBitmapEncoder()
            Case ".wdp"
                encoder = New WmpBitmapEncoder()
            Case Else
                Throw New NotSupportedException
        End Select

        SaveUsingEncoder(Visual, fileName, encoder)
    End Sub

    Private Sub SaveUsingEncoder(ByVal Visual1 As FrameworkElement, ByVal fileName As String, ByVal encoder As BitmapEncoder)
        Dim W = Math.Floor(Visual1.ActualWidth)
        Dim H = Math.Floor(Visual1.ActualHeight)
        Dim bitmap As New RenderTargetBitmap(W * DpiX / 96, H * DpiY / 96, DpiX, DpiY, Me.PixelFormat)
        Dim dv As New DrawingVisual()
        Using dc As DrawingContext = dv.RenderOpen()
            Dim vb As New VisualBrush(Visual1)
            dc.DrawRectangle(vb, Nothing, New Rect(0, 0, W, H))
        End Using
        bitmap.Render(dv)
        'bitmap.Render(Visual1)
        Dim frame As BitmapFrame = BitmapFrame.Create(bitmap)
        encoder.Frames.Add(frame)

        Using stream = IO.File.Create(fileName)
            encoder.Save(stream)
        End Using
    End Sub

End Class
