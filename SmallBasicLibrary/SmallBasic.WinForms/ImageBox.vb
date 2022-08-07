Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallBasic.Library.Internal.SmallBasicApplication
Imports System.Windows.Media.Imaging

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class ImageBox

        Private Shared Function GetImageBox(formName As String, imageBoxName As String) As Wpf.Image
            formName = formName.ToLower()
            If Not Forms._forms.ContainsKey(formName) Then
                Throw New Exception($"There is no form named `{formName}`.")
            End If

            Dim _controls = Forms._forms(formName)
            If imageBoxName = "" Then
                Throw New Exception($"imageBoxName can't be empty")
            End If

            Dim name = imageBoxName.ToLower()
            If Not _controls.ContainsKey(name) Then
                Throw New Exception($"There is no ImageBox named `{imageBoxName}` on form {formName}.")
            End If

            Dim imgBox = TryCast(_controls(name), Wpf.Image)
            If imgBox Is Nothing Then
                Throw New Exception($"{imageBoxName} is not a name of a ImageBox.")
            End If
            Return imgBox
        End Function

        ''' <summary>
        ''' Gets or sets the path of the image that is displayed in the control.
        ''' </summary>
        ''' <param name="formName"></param>
        ''' <param name="ImageBoxName"></param>
        ''' <returns></returns>
        <ExProperty>
        Public Shared Function GetFileName(formName As Primitive, ImageBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim img = GetImageBox(formName, ImageBoxName)
                        Dim src = CType(img.Source, BitmapImage)
                        If src IsNot Nothing Then
                            GetFileName = src.UriSource?.AbsolutePath
                        Else
                            GetFileName = ""
                        End If

                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ImageBoxName, "Image", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetFileName(formName As Primitive, ImageBoxName As Primitive, imageFile As Primitive)
            App.Invoke(
                Sub()
                    Try
                        If Not IO.Path.IsPathRooted(imageFile) Then
                            Dim path = Environment.GetCommandLineArgs()(0)
                            path = IO.Path.GetDirectoryName(path)
                            imageFile = IO.Path.Combine(path, imageFile)
                        End If
                        GetImageBox(formName, ImageBoxName).Source = New BitmapImage(New Uri(imageFile))
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ImageBoxName, "Text", imageFile, ex.Message)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' The x-pos of the control on its parent control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetLeft(formName As Primitive, ImageBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetLeft = Wpf.Canvas.GetLeft(GetImageBox(formName, ImageBoxName))
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ImageBoxName, "Left", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetLeft(formName As Primitive, ImageBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Wpf.Canvas.SetLeft(GetImageBox(formName, ImageBoxName), value)
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ImageBoxName, "Left", value, ex.Message)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The y-pos of the control on its parent control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetTop(formName As Primitive, ImageBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetTop = Wpf.Canvas.GetTop(GetImageBox(formName, ImageBoxName))
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ImageBoxName, "Top", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetTop(formName As Primitive, ImageBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Wpf.Canvas.SetTop(GetImageBox(formName, ImageBoxName), value)
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ImageBoxName, "Top", value, ex.Message)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The width of the control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetWidth(formName As Primitive, ImageBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetWidth = GetImageBox(formName, ImageBoxName).ActualWidth
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ImageBoxName, "Width", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetWidth(formName As Primitive, ImageBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetImageBox(formName, ImageBoxName).Width = value
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ImageBoxName, "Width", value, ex.Message)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The height of the control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetHeight(formName As Primitive, ImageBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetHeight = GetImageBox(formName, ImageBoxName).ActualHeight
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ImageBoxName, "Height", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetHeight(formName As Primitive, ImageBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetImageBox(formName, ImageBoxName).Height = value
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ImageBoxName, "Height", value, ex.Message)
                    End Try
                End Sub)
        End Sub

    End Class
End Namespace