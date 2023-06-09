
Imports System.Windows.Media.Imaging
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
            If fileNameOrUrl.IsEmpty Then Return ""

            Dim imageName = Shapes.GenerateNewName("ImageList")
            SmallBasicApplication.Invoke(
                Sub()
                    _savedImages(imageName) = LoadImageFromFile(fileNameOrUrl)
                End Sub)

            Return imageName
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
        ''' Gets the height of the stored image.
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

        Friend Shared Function GetBitmap(imageName As Primitive) As BitmapSource
            Dim value As BitmapSource = Nothing

            GraphicsWindow.Invoke(
                Sub()
                    If Not imageName.IsEmpty Then
                        _savedImages.TryGetValue(imageName, value)
                    End If
                    If value Is Nothing Then
                        value = LoadImageFromFile(imageName)
                    End If
                End Sub)

            Return value
        End Function

        Friend Shared Function LoadImageFromFile(fileNameOrUrl As Primitive) As BitmapImage
            Return CType(SmallBasicApplication.InvokeWithReturn(
                Function() As BitmapImage
                    Try
                        fileNameOrUrl = Environment.ExpandEnvironmentVariables(fileNameOrUrl)
                        Dim localFileName As String = Network.GetLocalFile(fileNameOrUrl)
                        If IO.Path.IsPathRooted(localFileName) Then
                            If localFileName.StartsWith("\") OrElse localFileName.StartsWith("/") Then
                                localFileName = IO.Path.Combine(Program.Directory, localFileName.TrimStart({"\"c, "/"c}))
                            End If
                        Else
                            localFileName = IO.Path.Combine(Program.Directory, localFileName)
                        End If
                        Return New BitmapImage(New Uri(localFileName, UriKind.RelativeOrAbsolute))

                    Catch
                    End Try
                    Return Nothing
                End Function), BitmapImage)
        End Function
    End Class
End Namespace
