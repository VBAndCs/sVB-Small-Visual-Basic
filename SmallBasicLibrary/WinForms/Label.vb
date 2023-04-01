Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows.Media.Imaging
Imports System.Windows
Imports System.Windows.Navigation
Imports System.Windows.Shapes

Namespace WinForms
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class Label
        ''' <summary>
        ''' Represents a Label control, that shows a text, a hyper link, an image or a graphic to the user.
        ''' Use the Text property to set the text displayed by the label. You can also use the Append methods to add formatted text to the label.
        ''' Use the AppendLink method to add a Hyper link to the label.
        ''' Use the Image property to load an image from a file and display it in the label.
        ''' Use the Add methods to add geometric shapes to the label.
        ''' You can use the form designer to add a label to the form by dragging it from the toolbox.
        ''' It is also possible to use the Form.AddLabel method to create a new label and add it to the form at runtime.
        ''' </summary>
        Private Shared Function GetLabel(labelName As String) As Wpf.Label
            Dim c = Control.GetControl(labelName)
            Dim lbl = TryCast(c, Wpf.Label)
            If lbl Is Nothing Then
                Throw New Exception($"{labelName} is not a name of a Label.")
            End If
            Return lbl
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
                            imageFile = IO.Path.Combine(Program.Directory, imageFile)
                        End If
                        GetLabel(labelName).Content = New Wpf.Image() With {
                            .Source = New BitmapImage(New Uri(imageFile))
                        }
                    Catch ex As Exception
                        Control.ReportError(labelName, "Image", imageFile, ex)
                    End Try
                End Sub)
        End Sub

        Friend Shared Function GetTextBlock(controlName As Primitive) As Wpf.TextBlock
            Dim cntrl = CType(Control.GetControl(controlName), Wpf.ContentControl)
            Dim content = cntrl.Content
            Dim tb = TryCast(content, Wpf.TextBlock)

            If tb Is Nothing Then
                tb = New Wpf.TextBlock()
                If content IsNot Nothing Then tb.Text = content.ToString()
                cntrl.Content = tb
                tb.TextWrapping = TextWrapping.Wrap
            End If

            Return tb
        End Function

        ''' <summary>
        ''' Adds the given text at the end of the current label with the given formats.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="fontName">The name of the font to apply on the text. Send an empty string to use the current label font</param>
        ''' <param name="fontSize">The font size of the text. Send 0 to use the current label font size</param>
        ''' <param name="isBold">True to use a bold font, False otherwise. Send an empty string to use the current label FontBold value.</param>
        ''' <param name="isItalic">True to use an italic font, False otherwise. Send an empty string to use the current label FontItalic value.</param>
        ''' <param name="isUnderlined">True to draw aline under the text, False otherwise. Send an empty string to use the current label Underlined value.</param>
        ''' <param name="foreColor">The color of the text. Send Colors.None to use the current label foreColor.</param>
        ''' <param name="backColor">The background color of the text. Send Colors.None to use the current label backColor.</param>
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
                           foreColor As Primitive,
                           backColor As Primitive,
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

                        If foreColor <> Colors.None Then
                            txtRun.Foreground = Color.GetBrush(foreColor)
                        End If

                        If backColor <> Colors.None Then
                            txtRun.Background = Color.GetBrush(backColor)
                        End If

                        Dim tb = GetTextBlock(labelName)
                        If url.AsString() = "" Then
                            tb.Inlines.Add(txtRun)

                        Else
                            Dim link As New Documents.Hyperlink(txtRun)
                            Try
                                url = GetAbsUrl(url)
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

        Public Shared Function GetAbsUrl(url As String) As String
            Try
                If IO.File.Exists(url) OrElse IO.Directory.Exists(url) Then
                    Return IO.Path.GetFullPath(url)
                ElseIf url.ToLower().StartsWith("www.") Then
                    Return "https://" & url
                End If
            Catch
            End Try

            Return Environment.ExpandEnvironmentVariables(url)
        End Function

        Private Shared Sub NavigateToLink(sender As Object, e As RequestNavigateEventArgs)
            Try
                Dim url = If(e.Uri.IsFile, e.Uri.LocalPath, e.Uri.AbsoluteUri)
                Process.Start(url)
            Catch ex As Exception
                Process.Start("about:blank")
                ReportError(ex.Message, ex)
            End Try
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with the given font name.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="fontName">The name of the font to apply on the text. Send an empty string to use the current label font</param>
        <ExMethod>
        Public Shared Sub AppendWithFontName(labelName As Primitive, text As Primitive, fontName As Primitive)
            AppendFormatted(labelName, text, fontName, 0, "", "", "", Colors.None, Colors.None, "")
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with the given font size.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="fontSize">The font size of the text. Send 0 to use the current label font size.</param>
        <ExMethod>
        Public Shared Sub AppendWithFontSize(labelName As Primitive, text As Primitive, fontSize As Primitive)
            AppendFormatted(labelName, text, "", fontSize, "", "", "", Colors.None, Colors.None, "")
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with the given font name and size.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="fontName">The name of the font to apply on the text. Send an empty string to use the current label font</param>
        ''' <param name="fontSize">The font size of the text. Send 0 to use the current label font size</param>
        <ExMethod>
        Public Shared Sub AppendWithFontNameAndSize(labelName As Primitive, text As Primitive, fontName As Primitive, fontSize As Primitive)
            AppendFormatted(labelName, text, fontName, fontSize, "", "", "", Colors.None, Colors.None, "")
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with a bold font.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        <ExMethod>
        Public Shared Sub AppendBold(labelName As Primitive, text As Primitive)
            AppendFormatted(labelName, text, "", 0, True, "", "", Colors.None, Colors.None, "")
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with an italic font.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        <ExMethod>
        Public Shared Sub AppendItalic(labelName As Primitive, text As Primitive)
            AppendFormatted(labelName, text, "", 0, "", True, "", Colors.None, Colors.None, "")
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with a bold and italic font.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        <ExMethod>
        Public Shared Sub AppendBoldItalic(labelName As Primitive, text As Primitive)
            AppendFormatted(labelName, text, "", 0, True, True, "", Colors.None, Colors.None, "")
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with a line drawn under it.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        <ExMethod>
        Public Shared Sub AppendUnderlined(labelName As Primitive, text As Primitive)
            AppendFormatted(labelName, text, "", 0, "", "", True, Colors.None, Colors.None, "")
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with the given font effects.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="isBold">True to use a bold font, False otherwise.</param>
        ''' <param name="isItalic">True to use an italic font, False otherwise.</param>
        ''' <param name="isUnderlined">True to draw aline under the text, False otherwise.</param>
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

        ''' <summary>
        ''' Adds the given text at the end of the current label with the given font properties.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="fontName">The name of the font to apply on the text. Send an empty string to use the current label font</param>
        ''' <param name="fontSize">The font size of the text. Send 0 to use the current label font size</param>
        ''' <param name="isBold">True to use a bold font, False otherwise.</param>
        ''' <param name="isItalic">True to use an italic font, False otherwise.</param>
        ''' <param name="isUnderlined">True to draw aline under the text, False otherwise.</param>
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

        ''' <summary>
        ''' Adds the given text at the end of the current label with the given for color.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="foreColor">The color of the text. Send Colors.None to use the current label foreColor.</param>
        <ExMethod>
        Public Shared Sub AppendWithForeColor(labelName As Primitive, text As Primitive, foreColor As Primitive)
            AppendFormatted(labelName, text, "", 0, "", "", "", foreColor, Colors.None, "")
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with the given back color.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="backColor">The background color of the text. Send Colors.None to use the current label backColor.</param>
        <ExMethod>
        Public Shared Sub AppendWithBackColor(labelName As Primitive, text As Primitive, backColor As Primitive)
            AppendFormatted(labelName, text, "", 0, "", "", "", Colors.None, backColor, "")
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with the given fore and back colors.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="foreColor">The color of the text. Send Colors.None to use the current label foreColor.</param>
        ''' <param name="backColor">The background color of the text. Send Colors.None to use the current label backColor.</param>
        <ExMethod>
        Public Shared Sub AppendWithColors(labelName As Primitive, text As Primitive, foreColor As Primitive, backColor As Primitive)
            AppendFormatted(labelName, text, "", 0, "", "", "", foreColor, backColor, "")
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label and formates it as a hyper link that opens the given url.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="url">The address to navigate to. Send empty string to view a normal text, otherwise the formated text will be viewed as a hyper link that opens the given url.</param>
        <ExMethod>
        Public Shared Sub AppendLink(labelName As Primitive, text As Primitive, url As Primitive)
            AppendFormatted(labelName, text, "", 0, "", "", "", Colors.None, Colors.None, url)
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with a bold font and formates it as a hyper link that opens the given url.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="url">The address to navigate to. Send empty string to view a normal text, otherwise the formated text will be viewed as a hyper link that opens the given url.</param>
        <ExMethod>
        Public Shared Sub AppendBoldLink(labelName As Primitive, text As Primitive, url As Primitive)
            AppendFormatted(labelName, text, "", 0, True, "", "", Colors.None, Colors.None, url)
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with an italic font and formates it as a hyper link that opens the given url.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="url">The address to navigate to. Send empty string to view a normal text, otherwise the formated text will be viewed as a hyper link that opens the given url.</param>
        <ExMethod>
        Public Shared Sub AppendItalicLink(labelName As Primitive, text As Primitive, url As Primitive)
            AppendFormatted(labelName, text, "", 0, "", True, "", Colors.None, Colors.None, url)
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label with a bold and italic font and formates it as a hyper link that opens the given url.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        ''' <param name="url">The address to navigate to. Use an empty string to view a normal text, otherwise the formated text will be viewed as a hyper link that opens the given url.</param>
        <ExMethod>
        Public Shared Sub AppendBoldItalicLink(labelName As Primitive, text As Primitive, url As Primitive)
            AppendFormatted(labelName, text, "", 0, True, True, "", Colors.None, Colors.None, url)
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
        <ExMethod>
        Public Shared Sub Append(labelName As Primitive, text As Primitive)
            AppendFormatted(labelName, text, "", 0, "", "", "", Colors.None, Colors.None, "")
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the current label then inserts a new line.
        ''' </summary>
        ''' <param name="text">The text to add to the label</param>
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

        ''' <summary>
        ''' Gets or sets whether or not to the text will be continue on the next line if it exceeds the width of the control.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetWordWrap(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetWordWrap = (GetTextBlock(controlName).TextWrapping = TextWrapping.Wrap)
                    Catch ex As Exception
                        Control.ReportError(controlName, "WordWrap", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetWordWrap(controlName As Primitive, Value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           GetTextBlock(controlName).TextWrapping = If(Value, TextWrapping.Wrap, TextWrapping.NoWrap)
                       Catch ex As Exception
                           Control.ReportPropertyError(controlName, "WordWrap", Value, ex)
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

        ''' <summary>
        ''' Draws a rectangle shape with the specified width and height on the current label..
        ''' Note that the label can contain only one shape, but you can add many shapes to the GeometricPath then use the Label.AddGeometricPath to add them as a combined shape to the label.
        ''' </summary>
        ''' <param name="width">The width of the rectangle shape.</param>
        ''' <param name="height">The height of the rectangle shape.</param>
        ''' <param name="penColor">The color used to draw the shape outline</param>
        ''' <param name="penWidth">The width of the shape outline</param>
        ''' <param name="brushColor">The color used to fill the shape</param>
        <ExMethod>
        Public Shared Sub AddRectangle(
                         labelName As Primitive,
                         width As Primitive,
                         height As Primitive,
                         penColor As Primitive,
                         penWidth As Primitive,
                         brushColor As Primitive
                   )

            GraphicsWindow.Invoke(
                Sub()
                    Try
                        Dim lbl = GetLabel(labelName)
                        lbl.Width = Double.NaN
                        lbl.Height = Double.NaN
                        lbl.Background = Nothing
                        lbl.Content = New Rectangle With {
                            .Width = width,
                            .Height = height,
                            .Fill = Color.GetBrush(brushColor),
                            .Stroke = Color.GetBrush(penColor),
                            .StrokeThickness = penWidth
                        }

                    Catch ex As Exception
                        Control.ReportSubError(labelName, "AddRectangle", ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Draws an ellipse shape with the specified width and height on the currentLabel.
        ''' Note that the label can contain only one shape, but you can add many shapes to the GeometricPath then use the Label.AddGeometricPath to add them as a combined shape to the label.
        ''' </summary>
        ''' <param name="width">The width of the ellipse shape.</param>
        ''' <param name="height">The height of the ellipse shape.</param>
        ''' <param name="penColor">The color used to draw the shape outline</param>
        ''' <param name="penWidth">The width of the shape outline</param>
        ''' <param name="brushColor">The color used to fill the shape</param>
        <ExMethod>
        Public Shared Sub AddEllipse(
                         labelName As Primitive,
                         width As Primitive,
                         height As Primitive,
                         penColor As Primitive,
                         penWidth As Primitive,
                         brushColor As Primitive
                   )

            App.Invoke(
                Sub()
                    Try
                        Dim lbl = GetLabel(labelName)
                        lbl.Width = Double.NaN
                        lbl.Height = Double.NaN
                        lbl.Background = Nothing
                        lbl.Content = New Ellipse With {
                              .Width = width,
                              .Height = height,
                              .Fill = Color.GetBrush(brushColor),
                              .Stroke = Color.GetBrush(penColor),
                              .StrokeThickness = penWidth
                        }

                    Catch ex As Exception
                        Control.ReportSubError(labelName, "AddEllipse", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Draws a triangle shape represented by the specified points on the current label.
        ''' Note that the label can contain only one shape, but you can add many shapes to the GeometricPath then use the Label.AddGeometricPath to add them as a combined shape to the label.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the first point.</param>
        ''' <param name="y1">The y co-ordinate of the first point.</param>
        ''' <param name="x2">The x co-ordinate of the second point.</param>
        ''' <param name="y2">The y co-ordinate of the second point.</param>
        ''' <param name="x3">The x co-ordinate of the third point.</param>
        ''' <param name="y3">The y co-ordinate of the third point.</param>
        ''' <param name="penColor">The color used to draw the shape outline</param>
        ''' <param name="penWidth">The width of the shape outline</param>
        ''' <param name="brushColor">The color used to fill the shape</param>
        <ExMethod>
        Public Shared Sub AddTriangle(
                         labelName As Primitive,
                         x1 As Primitive, y1 As Primitive,
                         x2 As Primitive, y2 As Primitive,
                         x3 As Primitive, y3 As Primitive,
                         penColor As Primitive,
                         penWidth As Primitive,
                         brushColor As Primitive
                   )

            App.Invoke(
                Sub()
                    Try
                        Dim lbl = GetLabel(labelName)
                        lbl.Width = Double.NaN
                        lbl.Height = Double.NaN
                        lbl.Background = Nothing
                        lbl.Content = New Polygon With {
                              .Points = New Media.PointCollection() From {
                                    New Point(x1, y1),
                                    New Point(x2, y2),
                                    New Point(x3, y3)
                              },
                              .Fill = Color.GetBrush(brushColor),
                              .Stroke = Color.GetBrush(penColor),
                              .StrokeThickness = penWidth
                        }

                    Catch ex As Exception
                        Control.ReportSubError(labelName, "AddTriangle", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Draws a polygon shape represented by the given points array on the current label.       
        ''' Note that the label can contain only one shape, but you can add many shapes to the GeometricPath then use the Label.AddGeometricPath to add them as a combined shape to the label.
        ''' </summary>
        ''' <param name="pointsArr">An array of points representing the heads of the polygn. Each item in this array is an array containing the x and y of the point.</param>
        ''' <param name="penColor">The color used to draw the shape outline</param>
        ''' <param name="penWidth">The width of the shape outline</param>
        ''' <param name="brushColor">The color used to fill the shape</param>
        <ExMethod>
        Public Shared Sub AddPolygon(
                         labelName As Primitive,
                         pointsArr As Primitive,
                         penColor As Primitive,
                         penWidth As Primitive,
                         brushColor As Primitive
                   )

            If pointsArr.IsEmpty OrElse Not pointsArr.IsArray Then Return

            App.Invoke(
                Sub()
                    Try
                        Dim Points As New Media.PointCollection()
                        For Each point In pointsArr._arrayMap.Values
                            Points.Add(New Point(point.Items(1), point.Items(2)))
                        Next

                        Dim lbl = GetLabel(labelName)
                        lbl.Width = Double.NaN
                        lbl.Height = Double.NaN
                        lbl.Background = Nothing
                        lbl.Content = New Polygon With {
                             .Points = Points,
                             .Fill = Color.GetBrush(brushColor),
                             .Stroke = Color.GetBrush(penColor),
                             .StrokeThickness = penWidth
                        }

                    Catch ex As Exception
                        Control.ReportSubError(labelName, "AddPolygon", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Draws a line between the specified two points on the current label.
        ''' Note that the label can contain only one shape, but you can add many shapes to the GeometricPath then use the Label.AddGeometricPath to add them as a combined shape to the label.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the first point.</param>
        ''' <param name="y1">The y co-ordinate of the first point.''' </param>
        ''' <param name="x2">The x co-ordinate of the second point.</param>
        ''' <param name="y2">The y co-ordinate of the second point.</param>
        ''' <param name="penColor">The color used to draw the shape outline</param>
        ''' <param name="penWidth">The width of the shape outline</param>
        <ExMethod>
        Public Shared Sub AddLine(
                         labelName As Primitive,
                          x1 As Primitive, y1 As Primitive,
                          x2 As Primitive, y2 As Primitive,
                         penColor As Primitive,
                         penWidth As Primitive
                   )

            App.Invoke(
                Sub()
                    Try
                        Dim lbl = GetLabel(labelName)
                        lbl.Width = Double.NaN
                        lbl.Height = Double.NaN
                        lbl.Background = Nothing
                        lbl.Content = New Line With {
                            .X1 = x1,
                            .Y1 = y1,
                            .X2 = x2,
                            .Y2 = y2,
                            .Stroke = Color.GetBrush(penColor),
                            .StrokeThickness = penWidth
                        }

                    Catch ex As Exception
                        Control.ReportSubError(labelName, "AddLine", ex)
                    End Try
                End Sub)
        End Sub

    End Class
End Namespace