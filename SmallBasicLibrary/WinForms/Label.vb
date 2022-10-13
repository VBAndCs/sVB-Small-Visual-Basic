Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows.Media.Imaging

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
                        GetText = GetLabel(labelName).Content.ToString()
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
                        Dim strValue = CStr(value)
                        GetLabel(labelName).Content = strValue
                    Catch ex As Exception
                        Control.RepotPropertyError(labelName, "Text", value, ex)
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
                        lbl.Content = path
                    Catch ex As Exception
                        Control.ReportSubError(labelName, "AddGeometricPath", ex)
                    End Try
                End Sub)
        End Sub

    End Class
End Namespace