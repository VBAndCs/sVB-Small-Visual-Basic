Imports System.IO
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' This class provides methods to interact with the desktop.
    ''' </summary>
    <SmallVisualBasicType>
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

        ''' <summary>
        ''' Displays the font dialog to allow the user to choose font name, size and other font properties.
        ''' </summary>
        ''' <param name="font">An array that  contains the initial properties to display in the dialog. You can get this arry from the font property of the control that you show the font dialog for.</param>
        ''' <returns>an array containing the font properties under the keys Name, Size, Bold, Italic, Underlined and Color, or returns an empty string "" if the user canceled the operation</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function ShowFontDialog(font As Primitive) As Primitive
            Dim key As New Primitive("Color")
            Dim wpfColor = WinForms.Color.FromString(font.Items(key))
            Dim color = System.Drawing.Color.FromArgb(wpfColor.A, wpfColor.R, wpfColor.G, wpfColor.B)

            key._stringValue = "Name"
            Dim family As New System.Drawing.FontFamily(font.Items(key))

            Dim style As New System.Drawing.FontStyle
            key._stringValue = "Bold"
            If font.Items(key) Then style = style Or System.Drawing.FontStyle.Bold

            key._stringValue = "Italic"
            If font.Items(key) Then style = style Or System.Drawing.FontStyle.Italic

            key._stringValue = "Underlined"
            If font.Items(key) Then style = style Or System.Drawing.FontStyle.Underline

            key._stringValue = "Size"
            Dim dlg As New Forms.FontDialog() With {
                .AllowScriptChange = True,
                .AllowSimulations = True,
                .FontMustExist = True,
                .MaxSize = 72,
                .MinSize = 1,
                .ShowColor = True,
                .ShowHelp = True,
                .ShowEffects = True,
                .Font = New System.Drawing.Font(family, font.Items(key), style),
                .Color = color
            }

            If dlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Dim f = dlg.Font
                font = New Primitive
                key._stringValue = "Name"
                font.Items(key) = New Primitive(f.FontFamily.Name)

                key._stringValue = "Size"
                font.Items(key) = f.Size

                key._stringValue = "Bold"
                font.Items(key) = (f.Style And System.Drawing.FontStyle.Bold) > 0

                key._stringValue = "Italic"
                font.Items(key) = (f.Style And System.Drawing.FontStyle.Italic) > 0

                key._stringValue = "Underlined"
                font.Items(key) = (f.Style And System.Drawing.FontStyle.Underline) > 0

                key._stringValue = "Color"
                font.Items(key) = WinForms.Color.FromARGB(dlg.Color.A, dlg.Color.R, dlg.Color.G, dlg.Color.B)
                Return font
            End If

            Return New Primitive("")
        End Function

        ''' <summary>
        ''' Gets an array containing the font names defined on the users system.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared ReadOnly Property FontNames() As Primitive
            Get
                Dim fonts As New Dictionary(Of Primitive, Primitive)
                Dim n = 1
                Dim _fontNames As New List(Of String)

                For Each family In Media.Fonts.SystemFontFamilies
                    _fontNames.Add(family.Source)
                Next

                _fontNames.Sort()
                For Each fontName In _fontNames
                    fonts(n) = New Primitive(fontName)
                    n += 1
                Next

                Return New Primitive() With {.ArrayMap = fonts}
            End Get
        End Property


        ''' <summary>
        ''' Copies the specified text to the Windows clipboard.
        ''' </summary>
        ''' <param name="text">The text to copy to the clipboard.</param>
        Public Shared Sub SetClipboardText(text As Primitive)
            SmallBasicApplication.Invoke(
                Sub() If Not text.IsEmpty Then Clipboard.SetText(text)
            )
        End Sub

        ''' <summary>
        ''' Gets the text currently stored in the Windows clipboard.
        ''' </summary>
        ''' <returns>The text from the clipboard, or an empty string if no text is available.</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetClipboardText() As Primitive
            Return New Primitive(SmallBasicApplication.InvokeWithReturn(
                Function()
                    If Clipboard.ContainsText() Then
                        Return Clipboard.GetText()
                    End If
                    Return ""
                End Function))
        End Function

        Friend Shared Sub CaptureWindow(visual As FrameworkElement, dpi As Double, fileName As String)
            If visual Is Nothing Then Return

            Dim w = visual.ActualWidth
            Dim h = visual.ActualHeight

            If w <= 0 OrElse h <= 0 Then Return

            Try
                ' High DPI scaling factor (e.g., 300 DPI instead of standard 96 DPI for crisp scaling)
                Dim dpiX = dpi
                Dim dpiY = dpi

                Dim scaleX = dpiX / 96.0
                Dim scaleY = dpiY / 96.0

                Dim bitmap As New RenderTargetBitmap(
                    CInt(w * scaleX), CInt(h * scaleY),
                    dpiX, dpiX, PixelFormats.Pbgra32)
                bitmap.Render(visual)

                If bitmap.CanFreeze Then
                    bitmap.Freeze()
                End If

                If String.IsNullOrEmpty(fileName) Then
                    ' Copy the rendered high-DPI image to the Windows clipboard safely
                    Clipboard.SetImage(bitmap)
                Else
                    fileName = Network.GetLocalFile(fileName)
                    ' Determine the appropriate encoder based on the file extension
                    Dim encoder As BitmapEncoder
                    Select Case IO.Path.GetExtension(fileName).ToLower()
                        Case ".bmp"
                            encoder = New BmpBitmapEncoder()
                        Case ".png"
                            encoder = New PngBitmapEncoder()
                        Case ".gif"
                            encoder = New GifBitmapEncoder()
                        Case ".jpg", ".jpeg"
                            encoder = New JpegBitmapEncoder()
                        Case ".tiff"
                            encoder = New TiffBitmapEncoder()
                        Case ".wdp"
                            encoder = New WmpBitmapEncoder()
                        Case Else
                            encoder = New PngBitmapEncoder()
                    End Select

                    encoder.Frames.Add(BitmapFrame.Create(bitmap))

                    Using stream As IO.Stream = IO.File.Create(fileName)
                        encoder.Save(stream)
                    End Using
                End If

            Catch ex As Exception
                WinForms.Control.ReportSubError(fileName, "SaveScreen", ex)
            End Try
        End Sub
    End Class
End Namespace
