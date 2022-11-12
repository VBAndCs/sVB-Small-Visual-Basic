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
        ''' <param name="font">The initial font to show its properties in the dialog.</param>
        ''' <returns>an array containing the font properties under the keys Name, Size, Bold, Italic, Underlined and Color, or returns an empty string "" if the user canceled the operation</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function ShowFontDialog(font As Primitive) As Primitive
            Dim wpfColor = WinForms.Color.FromString(font.Items("Color"))
            Dim color = Drawing.Color.FromArgb(wpfColor.A, wpfColor.R, wpfColor.G, wpfColor.B)

            Dim family As New Drawing.FontFamily(font.Items("Name"))
            Dim style As New Drawing.FontStyle
            If font.Items("Bold") Then style = style Or Drawing.FontStyle.Bold
            If font.Items("Italic") Then style = style Or Drawing.FontStyle.Italic
            If font.Items("Underlined") Then style = style Or Drawing.FontStyle.Underline

            Dim dlg As New Forms.FontDialog() With {
                .AllowScriptChange = True,
                .AllowSimulations = True,
                .FontMustExist = True,
                .MaxSize = 72,
                .MinSize = 1,
                .ShowColor = True,
                .ShowHelp = True,
                .ShowEffects = True,
                .Font = New Drawing.Font(family, font.Items("Size"), style),
                .Color = color
            }

            If dlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Dim f = dlg.Font
                font = New Primitive
                font.Items("Name") = f.FontFamily.Name
                font.Items("Size") = f.Size
                font.Items("Bold") = (f.Style And Drawing.FontStyle.Bold) > 0
                font.Items("Italic") = (f.Style And Drawing.FontStyle.Italic) > 0
                font.Items("Underlined") = (f.Style And Drawing.FontStyle.Underline) > 0
                font.Items("Color") = WinForms.Color.FromARGB(dlg.Color.A, dlg.Color.R, dlg.Color.G, dlg.Color.B)
                Return font
            End If

            Return ""
        End Function

        ''' <summary>
        ''' Gets an array containing the font names defined on the users system.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared ReadOnly Property FontNames() As Primitive
            Get
                Dim fonts As New Dictionary(Of Primitive, Primitive)
                Dim n = 1

                For Each family In Media.Fonts.SystemFontFamilies
                    fonts(n) = family.Source
                    n += 1
                Next

                Dim result As New Primitive
                result._arrayMap = fonts
                Return result
            End Get
        End Property
    End Class
End Namespace
