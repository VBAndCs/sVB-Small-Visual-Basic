Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows.Media.Imaging

Namespace WinForms
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class ImageBox

        Private Shared Function GetImageBox(imageBoxName As String) As Wpf.Image

            Dim c = Control.GetFrameworkElement(imageBoxName)
            Dim imgBox = TryCast(c, Wpf.Image)
            If imgBox Is Nothing Then
                Throw New Exception($"{imageBoxName} is not a name of a ImageBox.")
            End If
            Return imgBox
        End Function

        ''' <summary>
        ''' Gets or sets the path of the image that is displayed in the control.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetFileName(ImageBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim img = GetImageBox(ImageBoxName)
                        Dim src = CType(img.Source, BitmapImage)
                        If src IsNot Nothing Then
                            GetFileName = src.UriSource?.AbsolutePath
                        Else
                            GetFileName = ""
                        End If

                    Catch ex As Exception
                        Control.ReportError(ImageBoxName, "Image", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetFileName(ImageBoxName As Primitive, imageFile As Primitive)
            App.Invoke(
                Sub()
                    Try
                        If Not IO.Path.IsPathRooted(imageFile) Then
                            Dim path = Environment.GetCommandLineArgs()(0)
                            path = IO.Path.GetDirectoryName(path)
                            imageFile = IO.Path.Combine(path, imageFile)
                        End If
                        GetImageBox(ImageBoxName).Source = New BitmapImage(New Uri(imageFile))
                    Catch ex As Exception
                        Control.ReportError(ImageBoxName, "Text", imageFile, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' The x-pos of the control on its parent control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetLeft(ImageBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetLeft = Wpf.Canvas.GetLeft(GetImageBox(ImageBoxName))
                    Catch ex As Exception
                        Control.ReportError(ImageBoxName, "Left", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetLeft(ImageBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Wpf.Canvas.SetLeft(GetImageBox(ImageBoxName), value)
                    Catch ex As Exception
                        Control.ReportPropertyError(ImageBoxName, "Left", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The y-pos of the control on its parent control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetTop(ImageBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetTop = Wpf.Canvas.GetTop(GetImageBox(ImageBoxName))
                    Catch ex As Exception
                        Control.ReportError(ImageBoxName, "Top", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetTop(ImageBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Wpf.Canvas.SetTop(GetImageBox(ImageBoxName), value)
                    Catch ex As Exception
                        Control.ReportPropertyError(ImageBoxName, "Top", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The width of the control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetWidth(ImageBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetWidth = GetImageBox(ImageBoxName).ActualWidth
                    Catch ex As Exception
                        Control.ReportError(ImageBoxName, "Width", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetWidth(ImageBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetImageBox(ImageBoxName).Width = value
                    Catch ex As Exception
                        Control.ReportPropertyError(ImageBoxName, "Width", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The height of the control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetHeight(ImageBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetHeight = GetImageBox(ImageBoxName).ActualHeight
                    Catch ex As Exception
                        Control.ReportError(ImageBoxName, "Height", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetHeight(ImageBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetImageBox(ImageBoxName).Height = value
                    Catch ex As Exception
                        Control.ReportPropertyError(ImageBoxName, "Height", value, ex)
                    End Try
                End Sub)
        End Sub

    End Class
End Namespace