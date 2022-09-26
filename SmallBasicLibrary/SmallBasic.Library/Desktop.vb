Imports System.IO
Imports System.Windows
Imports System.Windows.Media.Imaging
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' This class provides methods to interact with the desktop.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Desktop
        ''' <summary>
        ''' Gets the screen width of the primary desktop.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared ReadOnly Property Width As Primitive
            Get
                Return SystemParameters.PrimaryScreenWidth
            End Get
        End Property

        ''' <summary>
        ''' Gets the screen height of the primary desktop.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared ReadOnly Property Height As Primitive
            Get
                Return SystemParameters.PrimaryScreenHeight
            End Get
        End Property

        ''' <summary>
        ''' Sets the specified picture as the desktop's wallpaper.  This file could be a local file or a network file or even an Internet URL.
        ''' </summary>
        ''' <param name="fileOrUrl">
        ''' The filename or URL of the picture.
        ''' </param>
        Public Shared Sub SetWallPaper(fileOrUrl As Primitive)
            If Not fileOrUrl.IsEmpty Then
                Dim lpvParam As String = ConvertToBitmap(fileOrUrl)
                NativeHelper.SystemParametersInfo(20, 0, lpvParam, 3)
            End If
        End Sub

        Friend Shared Function ConvertToBitmap(fileNameOrUrl As Primitive) As Primitive
            If fileNameOrUrl.IsEmpty Then
                Return fileNameOrUrl
            End If

            Dim localFile As String = Path.GetTempFileName()
            SmallBasicApplication.Invoke(
                Sub()
                    Dim bitmap = ImageList.LoadImageFromFile(fileNameOrUrl)
                    If bitmap IsNot Nothing Then
                        Dim bmpEncoder As New BmpBitmapEncoder
                        Dim item = BitmapFrame.Create(bitmap)
                        bmpEncoder.Frames.Add(item)
                        Using stream = IO.File.Create(localFile)
                            bmpEncoder.Save(stream)
                        End Using
                    End If
                End Sub)
            Return New Primitive(localFile)
        End Function
    End Class
End Namespace
