Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows.Media.Imaging
Imports System.Windows
Imports System.Windows.Navigation

Namespace WinForms
    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class Label

        Private Shared Function GetLabel(labelName As String) As Wpf.Label
            Dim c = Control.GetControl(labelName)
            Dim t = TryCast(c, Wpf.Label)
            If t Is Nothing Then
                Throw New Exception($"{labelName} is not a name of a Label.")
            End If
            Return t
        End Function

        ''' <summary>
        ''' Gets or sets the text that is displayed on the label
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetText(labelName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetText = GetTextBlock(labelName).Text
                    Catch ex As Exception
                        Control.ReportError(labelName, "Text", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetText(labelName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetTextBlock(labelName).Text = value.AsString()
                    Catch ex As Exception
                        Control.ReportPropertyError(labelName, "Text", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the path of the image that is displayed on the label
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetImage(labelName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim content = GetLabel(labelName).Content
                        Dim img = TryCast(content, Wpf.Image)
                        If img IsNot Nothing Then
                            GetImage = CType(img.Source, BitmapImage).UriSource.AbsolutePath
                        Else
                            GetImage = ""
                        End If

                    Catch ex As Exception
                        Control.ReportError(labelName, "Image", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetImage(labelName As Primitive, imageFile As Primitive)
            App.Invoke(
                Sub()
                    Try
                        If Not IO.Path.IsPathRooted(imageFile) Then
                            Dim path = Environment.GetCommandLineArgs()(0)
                            path = IO.Path.GetDirectoryName(path)
                            imageFile = IO.Path.Combine(path, imageFile)
                        End If
                        GetLabel(labelName).Content = New Wpf.Image() With {.Source = New BitmapImage(New Uri(imageFile))}
                    Catch ex As Exception
                        Control.ReportError(labelName, "Text", imageFile, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Draws the shape defined by GeometricPath on the label
        ''' </summary>
        ''' <param name="penColor">The color used to draw the shape outline</param>
        ''' <param name="penWidth">The width of the shape outline</param>
        ''' <param name="brushColor">The color used to fill the shape</param>
        <ExMethod>
        Public Shared Sub AddGeometricPath(
                         labelName As Primitive,
                         penColor As Primitive,
                         penWidth As Primitive,
                         brushColor As Primitive
                   )

            App.Invoke(
                Sub()
                    Try
                        Dim path = GeometricPath._path
                        path.Fill = Color.GetBrush(brushColor)
                        path.Stroke = Color.GetBrush(penColor)
                        path.StrokeThickness = penWidth

                        Dim lbl = GetLabel(labelName)
                        lbl.Width = Double.NaN
                        lbl.Height = Double.NaN
                        lbl.Background = Nothing
                        lbl.Content = path
                    Catch ex As Exception
                        Control.ReportSubError(labelName, "AddGeometricPath", ex)
                    End Try
                End Sub)
        End Sub

        Friend Shared Function GetTextBlock(controlName As Primitive) As Wpf.TextBlock
            Dim cntrl = CType(Control.GetControl(controlName), Wpf.ContentControl)
            Dim content = cntrl.Content
            Dim tb = TryCast(content, Wpf.TextBlock)
            If tb Is Nothing Then
                tb = New Wpf.TextBlock()
                tb.TextWrapping = TextWrapping.Wrap
                If content IsNot Nothing Then tb.Text = content.ToString()
                cntrl.Content = tb
            End If
            Return tb
        End Function

        ''' <summary>
        ''' Adds the given text at the end of the current label with the given formats.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="fontName">The name of the font to apply on the text. Send an empty string to use the current label font</param>
        ''' <param name="fontSize">The font size of the text. Send 0 to use the current label font size</param>
        ''' <param name="isBold">True to use a bold font, False otherwise.</param>
        ''' <param name="isItalic">True to use an italic font, False otherwise.</param>
        ''' <param name="isUnderlined">True to draw aline under the text, False otherwise.</param>
        ''' <param name="forecolor">The color of the text. Send Colors.None to use the current label forecolor.</param>
        ''' <param name="backcolor">The background color of the text. Send Colors.None to use the current label backcolor></param>
        ''' <param name="url">The address to navigate to. Send empty string to view a normal text, otherwise the formated text will be viewed as a hyper link that opens the given url.</param>
        <ExMethod>
        Public Shared Sub AppendFormatted(
                           labelName As Primitive,
                           text As Primitive,
                           fontName As Primitive,
                           fontSize As Primitive,
                           isBold As Primitive,
                           isItalic As Primitive,
                           isUnderlined As Primitive,
                           forecolor As Primitive,
                           backcolor As Primitive,
                           url As Primitive
                    )

            App.Invoke(
                Sub()
                    Try
                        Dim lbl = GetLabel(labelName)
                        lbl.Padding = New Thickness(5)

                        Dim txtRun As New Documents.Run()
                        txtRun.Text = text

                        If fontName.AsString() <> "" Then
                            txtRun.FontFamily = New Media.FontFamily(fontName)
                        End If

                        If fontSize > 0 Then
                            txtRun.FontSize = fontSize
                        End If

                        If isBold.AsString() <> "" Then
                            txtRun.FontWeight = If(isBold, FontWeights.Bold, FontWeights.Normal)
                        End If

                        If isItalic.AsString() <> "" Then
                            txtRun.FontStyle = If(isItalic, FontStyles.Italic, FontStyles.Normal)
                        End If

                        If isUnderlined.AsString() <> "" Then
                            If isUnderlined Then txtRun.TextDecorations = TextDecorations.Underline
                        End If

                        If forecolor <> Colors.None Then
                            txtRun.Foreground = Color.GetBrush(forecolor)
                        End If

                        If backcolor <> Colors.None Then
                            txtRun.Background = Color.GetBrush(backcolor)
                        End If

                        Dim tb = GetTextBlock(labelName)
                        If url.AsString() = "" Then
                            tb.Inlines.Add(txtRun)

                        Else
                            Dim link As New Documents.Hyperlink(txtRun)
                            Try
                                link.NavigateUri = New Uri(url)
                            Catch ex As Exception
                                link.NavigateUri = New Uri("about:blank")
                            End Try

                            AddHandler link.RequestNavigate, AddressOf NavigateToLink
                            tb.Inlines.Add(link)
                        End If

                    Catch ex As Exception
                        Control.ReportSubError(labelName, "AppendFormat", ex)
                    End Try
                End Sub)
        End Sub

        Private Shared Sub NavigateToLink(sender As Object, e As RequestNavigateEventArgs)
            Process.Start(e.Uri.AbsoluteUri)
        End Sub

        <ExMethod>
        Public Shared Sub AppendWithFontName(labelName As Primitive, text As Primitive, fontName As Primitive)
            AppendFormatted(labelName, text, fontName, 0, "", "", "", Colors.None, Colors.None, "")
        End Sub

        <ExMethod>
        Public Shared Sub AppendWithFontSize(labelName As Primitive, text As Primitive, fontSize As Primitive)
            AppendFormatted(labelName, text, "", fontSize, "", "", "", Colors.None, Colors.None, "")
        End Sub

        <ExMethod>
        Public Shared Sub AppendWithFontNameAndSize(labelName As Primitive, text As Primitive, fontName As Primitive, fontSize As Primitive)
            AppendFormatted(labelName, text, fontName, fontSize, "", "", "", Colors.None, Colors.None, "")
        End Sub

        <ExMethod>
        Public Shared Sub AppendBold(labelName As Primitive, text As Primitive)
            AppendFormatted(labelName, text, "", 0, True, "", "", Colors.None, Colors.None, "")
        End Sub


        <ExMethod>
        Public Shared Sub AppendItalic(labelName As Primitive, text As Primitive)
            AppendFormatted(labelName, text, "", 0, "", True, "", Colors.None, Colors.None, "")
        End Sub

        <ExMethod>
        Public Shared Sub AppendBoldItalic(labelName As Primitive, text As Primitive)
            AppendFormatted(labelName, text, "", 0, True, True, "", Colors.None, Colors.None, "")
        End Sub

        <ExMethod>
        Public Shared Sub AppendUnderlined(labelName As Primitive, text As Primitive)
            AppendFormatted(labelName, text, "", 0, "", "", True, Colors.None, Colors.None, "")
        End Sub

        <ExMethod>
        Public Shared Sub AppendWithFontEffects(
                           labelName As Primitive,
                           text As Primitive,
                           isBold As Primitive,
                           isItalic As Primitive,
                           isUnderlined As Primitive
                   )

            AppendFormatted(labelName, text, "", 0, isBold, isItalic, isUnderlined, Colors.None, Colors.None, "")
        End Sub

        <ExMethod>
        Public Shared Sub AppendWithFont(
                           labelName As Primitive,
                           text As Primitive,
                           fontName As Primitive,
                           fontSize As Primitive,
                           isBold As Primitive,
                           isItalic As Primitive,
                           isUnderlined As Primitive
                   )

            AppendFormatted(labelName, text, fontName, fontSize, isBold, isItalic, isUnderlined, Colors.None, Colors.None, "")
        End Sub

        <ExMethod>
        Public Shared Sub AppendWithForecolor(labelName As Primitive, text As Primitive, forecolor As Primitive)
            AppendFormatted(labelName, text, "", 0, "", "", "", forecolor, Colors.None, "")
        End Sub

        <ExMethod>
        Public Shared Sub AppendWithBackcolor(labelName As Primitive, text As Primitive, backcolor As Primitive)
            AppendFormatted(labelName, text, "", 0, "", "", "", Colors.None, backcolor, "")
        End Sub

        <ExMethod>
        Public Shared Sub AppendWithColors(labelName As Primitive, text As Primitive, forecolor As Primitive, backcolor As Primitive)
            AppendFormatted(labelName, text, "", 0, "", "", "", forecolor, backcolor, "")
        End Sub

        <ExMethod>
        Public Shared Sub AppendLink(labelName As Primitive, text As Primitive, url As Primitive)
            AppendFormatted(labelName, text, "", 0, "", "", "", Colors.None, Colors.None, url)
        End Sub

        <ExMethod>
        Public Shared Sub AppendBoldLink(labelName As Primitive, text As Primitive, url As Primitive)
            AppendFormatted(labelName, text, "", 0, True, "", "", Colors.None, Colors.None, url)
        End Sub

        <ExMethod>
        Public Shared Sub AppendItalicLink(labelName As Primitive, text As Primitive, url As Primitive)
            AppendFormatted(labelName, text, "", 0, "", True, "", Colors.None, Colors.None, url)
        End Sub

        <ExMethod>
        Public Shared Sub AppendBoldItalicLink(labelName As Primitive, text As Primitive, url As Primitive)
            AppendFormatted(labelName, text, "", 0, True, True, "", Colors.None, Colors.None, url)
        End Sub

        <ExMethod>
        Public Shared Sub Append(labelName As Primitive, text As Primitive)
            AppendFormatted(labelName, text, "", 0, "", "", "", Colors.None, Colors.None, "")
        End Sub

        <ExMethod>
        Public Shared Sub AppendLine(labelName As Primitive, text As Primitive)
            AppendFormatted(labelName, text + vbCrLf, "", 0, "", "", "", Colors.None, Colors.None, "")
        End Sub


        ''' <summary>
        ''' Gets or sets whether or not to draw a line under the text.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetUnderlined(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetUnderlined = (GetTextBlock(controlName).TextDecorations Is TextDecorations.Underline)
                    Catch ex As Exception
                        Control.ReportError(controlName, "Underlined", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetUnderlined(controlName As Primitive, Value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           GetTextBlock(controlName).TextDecorations = If(CBool(Value), TextDecorations.Underline, Nothing)
                       Catch ex As Exception
                           Control.ReportPropertyError(controlName, "Underlined", Value, ex)
                       End Try
                   End Sub)
        End Sub


    End Class
End Namespace