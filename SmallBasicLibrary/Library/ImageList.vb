
Imports System.Drawing
Imports System.Windows
Imports System.Windows.Media.Imaging
Imports System.Windows.Media.Media3D
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' This class helps to load and store images in memory.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class ImageList
        Private Shared _savedImages As New Dictionary(Of String, BitmapSource)

        ''' <summary>
        ''' Loads an image from a file or the Internet into memory.
        ''' </summary>
        ''' <param name="fileNameOrUrl">
        ''' The file name to load the image from.  This could be a local file or a URL to the Internet location.
        ''' </param>
        ''' <returns>
        ''' Returns the name of the image that was loaded.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function LoadImage(fileNameOrUrl As Primitive) As Primitive
            If fileNameOrUrl.IsEmpty Then Return New Primitive("")

            Dim imageName = Shapes.GenerateNewName("ImageList")
            SmallBasicApplication.Invoke(
                Sub()
                    _savedImages(imageName) = LoadImageFromFile(fileNameOrUrl)
                End Sub)

            Return New Primitive(imageName)
        End Function

        ''' <summary>
        ''' Gets the pixel width of the stored image.
        ''' </summary>
        ''' <param name="imageName">
        ''' The name of the image in memory.
        ''' </param>
        ''' <returns>
        ''' The width of the specified image.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetWidthOfImage(imageName As Primitive) As Primitive
            Return GraphicsWindow.InvokeWithReturn(
                Function() As Primitive
                    Try
                        Dim image = GetBitmap(imageName)
                        If image Is Nothing Then Return New Primitive(0)
                        Return New Primitive(image.PixelWidth)
                    Catch ex As Exception
                    End Try

                    Return New Primitive(0)
                End Function)

            Return 0
        End Function

        ''' <summary>
        ''' Gets the pixel height of the stored image.
        ''' </summary>
        ''' <param name="imageName">
        ''' The name of the image in memory.
        ''' </param>
        ''' <returns>
        ''' The height of the specified image.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetHeightOfImage(imageName As Primitive) As Primitive
            Return GraphicsWindow.InvokeWithReturn(
                 Function() As Primitive
                     Try
                         Dim image = GetBitmap(imageName)
                         If image Is Nothing Then Return New Primitive(0)
                         Return New Primitive(image.PixelHeight)
                     Catch ex As Exception
                     End Try

                     Return New Primitive(0)
                 End Function)
            Return 0
        End Function

        ''' <summary>
        ''' Gets the width of the stored image when displayed of the window without scaling.
        ''' </summary>
        ''' <param name="imageName">
        ''' The name of the image in memory.
        ''' </param>
        ''' <returns>
        ''' The width of the specified image.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetActualWidth(imageName As Primitive) As Primitive
            Return GraphicsWindow.InvokeWithReturn(
                Function() As Primitive
                    Try
                        Dim image = GetBitmap(imageName)
                        If image Is Nothing Then Return New Primitive(0)
                        Return New Primitive(image.Width)
                    Catch ex As Exception
                    End Try

                    Return New Primitive(0)
                End Function)

            Return 0
        End Function

        ''' <summary>
        ''' Gets the height of the stored image when displayed on the window without scaling.
        ''' </summary>
        ''' <param name="imageName">
        ''' The name of the image in memory.
        ''' </param>
        ''' <returns>
        ''' The height of the specified image.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetActualHeight(imageName As Primitive) As Primitive
            Return GraphicsWindow.InvokeWithReturn(
                 Function() As Primitive
                     Try
                         Dim image = GetBitmap(imageName)
                         If image Is Nothing Then Return New Primitive(0)
                         Return New Primitive(image.Height)
                     Catch ex As Exception
                     End Try

                     Return New Primitive(0)
                 End Function)
            Return 0
        End Function


        Friend Shared Function GetBitmap(imageName As Primitive) As BitmapSource
            Dim value As BitmapSource = Nothing

            GraphicsWindow.Invoke(
                Sub()
                    If imageName.IsEmpty Then
                        value = Nothing
                    ElseIf Not _savedImages.TryGetValue(imageName, value) Then
                        value = LoadImageFromFile(imageName)
                    End If
                End Sub)

            Return value
        End Function

        Friend Shared Function LoadImageFromFile(fileNameOrUrl As Primitive) As BitmapImage
            Return CType(SmallBasicApplication.InvokeWithReturn(
                Function() As BitmapImage
                    Try
                        fileNameOrUrl = New Primitive(Environment.ExpandEnvironmentVariables(fileNameOrUrl))
                        Dim localFileName As String = Network.GetLocalFile(fileNameOrUrl)
                        Return New BitmapImage(New Uri(localFileName, UriKind.RelativeOrAbsolute))

                    Catch ex As Exception
                    End Try

                    Return Nothing
                End Function), BitmapImage)
        End Function

        Public Shared Function CreateCollisionMask(bitmap As BitmapSource) As Boolean(,)
            Dim wb As New WriteableBitmap(bitmap)
            Dim w = wb.PixelWidth
            Dim h = wb.PixelHeight
            Dim n = wb.Format.BitsPerPixel \ 8
            Dim stride = w * n
            Dim pixelData(stride * h - 1) As Byte
            wb.CopyPixels(pixelData, stride, 0)
            Dim mask(w - 1, h - 1) As Boolean

            For y = 0 To h - 1
                Dim offset = y * stride
                For x = 0 To w - 1
                    mask(x, y) = pixelData(offset + x * n + 3) > 0
                Next
            Next

            Return mask
        End Function

        Private Shared CollisionMasks As New Dictionary(Of String, Boolean(,))()
        Private Shared DpiX As Double
        Private Shared DpiY As Double

        Shared Sub New()
            Dim matrix = GetDpiMatrix()
            DpiX = matrix.M11
            DpiY = matrix.M22
        End Sub

        Shared Function GetDpiMatrix() As Media.Matrix
            SmallBasicApplication.Invoke(
                Sub()
                    Dim wind = Application.Current.MainWindow
                    If wind Is Nothing Then
                        GraphicsWindow.VerifyAccess()
                        wind = GraphicsWindow._window
                    End If
                    Dim source = PresentationSource.FromVisual(wind)
                    GetDpiMatrix = source.CompositionTarget.TransformToDevice
                End Sub)
        End Function

        ''' <summary>
        ''' Investigates the collision of the two given images. They must be loaded in the ImageList.
        ''' You must also display the 2 images in full size without any scaling.
        ''' </summary>
        ''' <param name="x1">the left position of the first image</param>
        ''' <param name="y1">the top position of the first image</param>
        ''' <param name="imageName1">the name of the first image in the image list</param>
        ''' <param name="x2">the left position of the second image</param>
        ''' <param name="y2">the top position of the second image</param>
        ''' <param name="imageName2">the name of the second image in the image list</param>
        ''' <returns>Flase if any image doesn't exist in tghe ImagesList, or the two images are not collidibng. 
        ''' True otherwise.</returns>
        Public Shared Function Collide(
                    x1 As Primitive,
                    y1 As Primitive,
                    imageName1 As Primitive,
                    x2 As Primitive,
                    y2 As Primitive,
                    imageName2 As Primitive) As Primitive

            Collide = New Primitive(False)
            SmallBasicApplication.Invoke(
                Sub()
                    Dim key1 = imageName1.AsString()
                    If key1 = "" Then Return

                    Dim key2 = imageName2.AsString()
                    If key2 = "" Then Return

                    Dim image1 = GetBitmap(key1)
                    If image1 Is Nothing Then Return

                    Dim image2 = GetBitmap(key2)
                    If image2 Is Nothing Then Return

                    Dim left1 = CDbl(x1) / DpiX
                    Dim left2 = CDbl(x2) / DpiX
                    Dim top1 = CDbl(y1) / DpiY
                    Dim top2 = CDbl(y2) / DpiY

                    Dim width1 = image1.PixelWidth
                    Dim width2 = image2.PixelWidth
                    Dim height1 = image1.PixelHeight
                    Dim height2 = image2.PixelHeight

                    If left1 > left2 + width2 OrElse
                        left1 + width1 < left2 OrElse
                        top1 > top2 + height2 OrElse
                        top1 + height1 < top2 Then Return

                    Dim mask1, mask2 As Boolean(,)
                    If CollisionMasks.ContainsKey(key1) Then
                        mask1 = CollisionMasks(key1)
                    Else
                        mask1 = CreateCollisionMask(image1)
                        CollisionMasks.Add(key1, mask1)
                    End If

                    If CollisionMasks.ContainsKey(key2) Then
                        mask2 = CollisionMasks(key2)
                    Else
                        mask2 = CreateCollisionMask(image2)
                        CollisionMasks.Add(key2, mask2)
                    End If

                    Dim dx As Integer = left1 - left2
                    Dim dy As Integer = top1 - top2

                    For x = 0 To width1 - 1
                        Dim m2X = x + dx
                        If m2X >= width2 Then Return
                        If m2X < 0 Then Continue For

                        For y = 0 To height1 - 1
                            Dim m2Y = y + dy
                            If m2Y >= height2 Then Exit For
                            If m2Y < 0 OrElse Not mask1(x, y) Then Continue For

                            If mask2(m2X, m2Y) Then
                                Collide = New Primitive(True)
                                Return
                            End If
                        Next
                    Next
                End Sub)
        End Function

    End Class
End Namespace
