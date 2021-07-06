' To configure or remove Option's included in result, go to Options/Advanced Options...
Option Compare Text
Option Explicit On
Option Infer On
Option Strict On
Imports System.IO
Imports System.Windows
Imports System.Windows.Media.Imaging
Imports SmallBasicLibrary.Microsoft.SmallBasic.Library.Internal

Namespace Microsoft.SmallBasic.Library
    ''' <summary>
    ''' This class provides methods to interact with the desktop.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Desktop
        ''' <summary>
        ''' Gets the screen width of the primary desktop.
        ''' </summary>
        Public Shared ReadOnly Property Width As Primitive
            Get
                Return SystemParameters.PrimaryScreenWidth
            End Get
        End Property

        ''' <summary>
        ''' Gets the screen height of the primary desktop.
        ''' </summary>
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
            SmallBasicApplication.Invoke(Sub()
                                             Dim bitmapImage1 As BitmapImage = ImageList.LoadImageFromFile(fileNameOrUrl)
                                             If bitmapImage1 IsNot Nothing Then
                                                 Dim bmpBitmapEncoder1 As New BmpBitmapEncoder
                                                 Dim item As BitmapFrame = BitmapFrame.Create(bitmapImage1)
                                                 bmpBitmapEncoder1.Frames.Add(item)
                                                 Using stream1 As Stream = IO.File.Create(localFile)
                                                     bmpBitmapEncoder1.Save(stream1)
                                                 End Using
                                             End If
                                         End Sub)
            Return New Primitive(localFile)
        End Function
    End Class
End Namespace
