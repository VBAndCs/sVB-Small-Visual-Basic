Imports System.IO
Imports System.Windows
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
            Dim color = Drawing.Color.FromArgb(wpfColor.A, wpfColor.R, wpfColor.G, wpfColor.B)

            key._stringValue = "Name"
            Dim family As New Drawing.FontFamily(font.Items(key))

            Dim style As New Drawing.FontStyle
            key._stringValue = "Bold"
            If font.Items(key) Then style = style Or Drawing.FontStyle.Bold

            key._stringValue = "Italic"
            If font.Items(key) Then style = style Or Drawing.FontStyle.Italic

            key._stringValue = "Underlined"
            If font.Items(key) Then style = style Or Drawing.FontStyle.Underline

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
                .Font = New Drawing.Font(family, font.Items(key), style),
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
                font.Items(key) = (f.Style And Drawing.FontStyle.Bold) > 0

                key._stringValue = "Italic"
                font.Items(key) = (f.Style And Drawing.FontStyle.Italic) > 0

                key._stringValue = "Underlined"
                font.Items(key) = (f.Style And Drawing.FontStyle.Underline) > 0

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
    End Class
End Namespace
