Imports System.Windows.Markup
Imports System.Windows.Threading
Imports Microsoft.SmallBasic.Library
Imports Microsoft.SmallBasic.Library.Internal
Imports Wpf = System.Windows.Controls
Imports ControlsDictionay = System.Collections.Generic.Dictionary(Of String, System.Windows.Controls.Control)
Imports System.Windows.Controls
Imports System.Windows
Imports System.ComponentModel

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class Forms

        Friend Shared _forms As New Dictionary(Of String, ControlsDictionay)

        Shared Function GetForm(name As String) As Window
            name = name.ToLower()
            If Not _forms.ContainsKey(name) Then Return Nothing

            Dim wnd = _forms(name)(name)
            If wnd Is Nothing Then
                Throw New ArgumentException($"`{name}` is not a form or it is closed")
            End If

            Return wnd
        End Function

        Public Shared Function GetForms() As Primitive
            Dim map = New Dictionary(Of Primitive, Primitive)
            Dim num = 1
            For Each key In _forms.Keys
                map(num) = key
                num += 1
            Next
            Return Primitive.ConvertFromMap(map)
        End Function

        Public Shared Property AppPath As Primitive

        Private Shared _syncLock As New Object

        Public Shared Function LoadForm(formName As Primitive, xamlPath As Primitive) As Primitive
            Dim xamlContent As String = $"<Canvas Name=""{formName}"" Width=""700"" Height=""500"" xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""/>"

            Dim name = CStr(formName).ToLower()
            If name = "" Then
                Throw New ArgumentException("Form name can't be an empty string.")
            End If

            If _forms.ContainsKey(name) Then Return name

            SyncLock _syncLock
                SmallBasicApplication.Invoke(
                    Sub()
                        Dim canvas As Canvas
                        Dim xaml = CStr(xamlPath)

                        If xaml = "" Then
                            canvas = LoadContent(name & ".xaml")
                        ElseIf xaml.StartsWith("<") Then
                            canvas = XamlReader.Load(Xml.XmlReader.Create(New IO.StringReader(xaml)))
                        Else
                            canvas = LoadContent(xamlPath)
                        End If

                        Dim wnd As New Window() With {
                           .SizeToContent = SizeToContent.WidthAndHeight,
                           .WindowStartupLocation = WindowStartupLocation.CenterScreen,
                           .Name = name,
                           .Content = canvas
                       }

                        AddHandler wnd.Closing, AddressOf Form_Closing

                        Dim _controls = New ControlsDictionay()
                        _controls(name) = wnd
                        _forms(name) = _controls

                        ' Add control names:
                        Dim controls = canvas.GetChildren().ToList()
                        For n = controls.Count - 1 To 0 Step -1
                            Dim ui = controls(n)
                            Dim controlName = Automation.AutomationProperties.GetName(ui)

                            If controlName = "" Then
                                Dim fw = TryCast(ui, FrameworkElement)
                                controlName = fw.Name
                                If controlName = "" Then Continue For
                            End If

                            If TypeOf ui Is Wpf.Control Then
                                _controls(controlName.ToLower()) = ui
                                CType(ui, Wpf.Control).Name = controlName
                            Else
                                Dim left = Canvas.GetLeft(ui)
                                Dim top = Canvas.GetTop(ui)
                                canvas.Children.Remove(ui)
                                Dim lb As New Wpf.Label() With {
                                    .Content = ui,
                                    .Name = controlName
                              }
                                canvas.Children.Add(lb)
                                Canvas.SetLeft(lb, left)
                                Canvas.SetTop(lb, top)
                                _controls(controlName.ToLower()) = lb
                            End If

                            SetControlText(ui, GetControlText(ui))
                        Next

                    End Sub)

            End SyncLock

            ' Ensure Keyboard Module is loaded, 
            ' to create a global hanler for the PreviewKeyDown event
            Dim __ = Keyboard.LastKey

            Return name
        End Function

        Private Shared Sub SetControlText(control As UIElement, value As String)
            Try
                CObj(control).Text = value
            Catch
                Try
                    Dim x = CObj(control).Content
                    If x Is Nothing OrElse TypeOf x Is String Then
                        CObj(control).Content = value
                    End If
                Catch
                    ' TODO: Add a label to hold the text
                End Try

            End Try
        End Sub

        Public Shared Function GetControlText(control As UIElement) As String
            Try
                Return CObj(control).Text
            Catch
                Try
                    Dim x = CObj(control).Content
                    If x IsNot Nothing AndAlso TypeOf x Is String Then
                        Return CStr(CObj(control).Content)
                    End If
                Catch
                End Try
            End Try

            Return Automation.AutomationProperties.GetHelpText(control)
        End Function


        Public Shared Sub AddForm(formName As Primitive)
            Dim frm = GetForm(formName)
            If frm Is Nothing Then
                Dim xaml As String = $"<Canvas Name=""{formName}"" Width=""700"" Height=""500"" xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""/>"
                LoadForm(formName, xaml)
            End If

        End Sub


        Private Shared Sub Form_Closing(sender As Object, e As CancelEventArgs)
            Dim win = CType(sender, Window)
            Dim formName = win.Name
            Dim CodeFilePath = GetSetting("sVb", "Designer", "CodeFilePath", formName & ".exe")

            Try
                ' If the form is created from code, save its design to .xml file
                Dim newXamlPath = CodeFilePath.Substring(0, CodeFilePath.Length - 4) & ".xaml"

                If Not IO.File.Exists(newXamlPath) Then
                    Dim canvas = win.Content
                    IO.File.WriteAllText(newXamlPath, XamlWriter.Save(canvas))
                End If
            Finally
                _forms.Remove(formName)
                If _forms.Count = 0 AndAlso Not TextWindow._windowVisible AndAlso Not GraphicsWindow._windowVisible Then
                    SmallBasicApplication.End()
                End If
            End Try

        End Sub

        Private Shared Function LoadContent(xamlPath As String) As Canvas
            If IO.Path.GetPathRoot(xamlPath) = "" Then
                If AppPath.ToString() <> "" Then
                    xamlPath = IO.Path.Combine(AppPath, xamlPath)
                Else
                    Dim d = AppDomain.CurrentDomain.BaseDirectory
                    Dim xamlPath2 = IO.Path.Combine(d, xamlPath)
                    If IO.File.Exists(xamlPath2) Then
                        xamlPath = xamlPath2
                    End If
                End If
            End If


            Dim stream = IO.File.Open(xamlPath, IO.FileMode.Open)
            Dim canvas As Canvas = Nothing
            Try
                canvas = CType(XamlReader.Load(stream), Canvas)
            Catch
            End Try
            Return canvas
        End Function

        Public Shared Sub ShowMessage(message As Primitive, title As Primitive)
            SmallBasicApplication.Invoke(Sub() System.Windows.MessageBox.Show(message.ToString(), title.ToString()))
        End Sub
    End Class
End Namespace