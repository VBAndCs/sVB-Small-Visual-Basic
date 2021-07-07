
Imports System.Windows.Media.Imaging
Imports Microsoft.SmallBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' This class helps to load and store images in memory.
    ''' </summary>
    <SmallBasicType>
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
        Public Shared Function LoadImage(fileNameOrUrl As Primitive) As Primitive
            If fileNameOrUrl.IsEmpty Then
                Return ""
            End If
            Dim text As String = Shapes.GenerateNewName("ImageList")
            _savedImages(text) = LoadImageFromFile(fileNameOrUrl)
            Return text
        End Function

        ''' <summary>
        ''' Gets the width of the stored image.
        ''' </summary>
        ''' <param name="imageName">
        ''' The name of the image in memory.
        ''' </param>
        ''' <returns>
        ''' The width of the specified image.
        ''' </returns>
        Public Shared Function GetWidthOfImage(imageName As Primitive) As Primitive
            Dim image As BitmapSource = GetBitmap(imageName)
            If image Is Nothing Then
                Return 0
            End If

            Return CType(GraphicsWindow.InvokeWithReturn(Function() CType(image.PixelWidth, Primitive)), Primitive)
        End Function

        ''' <summary>
        ''' Gets the height of the stored image.
        ''' </summary>
        ''' <param name="imageName">
        ''' The name of the image in memory.
        ''' </param>
        ''' <returns>
        ''' The height of the specified image.
        ''' </returns>
        Public Shared Function GetHeightOfImage(imageName As Primitive) As Primitive
            Dim image As BitmapSource = GetBitmap(imageName)
            If image Is Nothing Then
                Return 0
            End If

            Return CType(GraphicsWindow.InvokeWithReturn(Function() CType(image.PixelHeight, Primitive)), Primitive)
        End Function

        Friend Shared Function GetBitmap(imageName As Primitive) As BitmapSource
            Dim value As BitmapSource = Nothing
            If Not imageName.IsEmpty Then
                _savedImages.TryGetValue(imageName, value)
            End If
            If value Is Nothing Then
                value = LoadImageFromFile(imageName)
            End If

            Return value
        End Function

        Friend Shared Function LoadImageFromFile(fileNameOrUrl As Primitive) As BitmapImage
            Dim localFileName As Primitive = Network.GetLocalFile(fileNameOrUrl)
            Return CType(SmallBasicApplication.InvokeWithReturn(
                Function() As BitmapImage
                    Try
                        Return New BitmapImage(New Uri(localFileName))
                    Catch
                    End Try
                    Return Nothing
                End Function), BitmapImage)
        End Function
    End Class
End Namespace
