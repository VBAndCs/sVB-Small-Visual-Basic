Imports System.Windows.Markup
Imports System.Windows.Threading
Imports Microsoft.SmallBasic.Library
Imports Microsoft.SmallBasic.Library.Internal
Imports Wpf = System.Windows.Controls
Imports ControlsDictionay = System.Collections.Generic.Dictionary(Of String, System.Windows.FrameworkElement)
Imports System.Windows.Controls
Imports System.Windows
Imports System.ComponentModel

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class Forms

        Friend Shared _forms As New ControlsDictionay
        Friend Shared _controls As New ControlsDictionay

        Friend Shared Function GetForm(name As String) As Window
            name = name.ToLower()
            Dim wnd As Window
            Dim dictionary = If(name.Contains("."), _controls, _forms)

            If Not dictionary.ContainsKey(name) Then Return Nothing
            wnd = dictionary(name)

            If wnd Is Nothing Then
                ReportError($"`{name}` is not a form or it is closed", Nothing)
            End If

            Return wnd
        End Function

        ''' <summary>
        ''' Returns an array containing the names of all forms you created.
        ''' </summary>
        ''' <returns></returns>
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

        ''' <summary>
        ''' Loads a form from its xaml file.
        ''' </summary>
        ''' <param name="formName">the name of the form</param>
        ''' <param name="xamlPath">the path pf the xaml file that contains the form design</param>
        ''' <returns>The name of the form</returns>
        Public Shared Function LoadForm(formName As Primitive, xamlPath As Primitive) As Primitive
            Dim form_Name = CStr(formName).ToLower()
            If form_Name = "" Then
                Dim ex As New ArgumentException("Form name can't be an empty string.")
                ReportError(ex.Message, ex)
            End If

            If _forms.ContainsKey(form_Name) Then Return form_Name

            SyncLock _syncLock
                SmallBasicApplication.Invoke(
                    Sub()
                        Try
                            Dim canvas As Canvas
                            Dim xaml = CStr(xamlPath)

                            If xaml = "" Then
                                canvas = LoadContent(form_Name & ".xaml")
                            ElseIf xaml.StartsWith("<") Then
                                canvas = XamlReader.Load(Xml.XmlReader.Create(New IO.StringReader(xaml)))
                            Else
                                canvas = LoadContent(xamlPath)
                            End If

                            Dim wnd As New Window() With {
                                   .SizeToContent = SizeToContent.WidthAndHeight,
                                   .WindowStartupLocation = WindowStartupLocation.CenterScreen,
                                   .Name = form_Name,
                                   .Title = Automation.AutomationProperties.GetHelpText(canvas),
                                   .Content = canvas,
                                   .ResizeMode = ResizeMode.CanMinimize
                            }

                            AddHandler wnd.Closing, AddressOf Form_Closing
                            AddHandler wnd.Closed, AddressOf Form_Closed

                            _forms(form_Name) = wnd
                            _controls(form_Name & "." & form_Name) = wnd

                            ' Add control names:
                            Dim controls = canvas.GetChildren().ToList()
                            For n = controls.Count - 1 To 0 Step -1
                                Dim ui = controls(n)
                                Dim fw = TryCast(ui, FrameworkElement)
                                Dim controlName = Automation.AutomationProperties.GetName(ui)

                                If controlName = "" Then
                                    controlName = fw.Name
                                    If controlName = "" Then Continue For
                                End If

                                If TypeOf ui Is Wpf.Control Then
                                    _controls(form_Name & "." & controlName.ToLower()) = ui
                                    CType(ui, Wpf.Control).Name = controlName
                                Else
                                    Dim left = Canvas.GetLeft(ui)
                                    Dim top = Canvas.GetTop(ui)
                                    canvas.Children.Remove(ui)
                                    Dim lb As New Wpf.Label() With {
                                            .Name = controlName,
                                            .Width = fw.Width,
                                            .Height = fw.Height,
                                            .HorizontalContentAlignment = HorizontalAlignment.Stretch,
                                            .VerticalContentAlignment = VerticalAlignment.Stretch,
                                            .Content = ui,
                                            .Background = Nothing
                                    }
                                    fw.ClearValue(FrameworkElement.WidthProperty)
                                    fw.ClearValue(FrameworkElement.HeightProperty)

                                    canvas.Children.Add(lb)
                                    Canvas.SetLeft(lb, left)
                                    Canvas.SetTop(lb, top)
                                    _controls(form_Name & "." & controlName.ToLower()) = lb
                                End If

                                SetControlText(ui, GetControlText(ui))
                            Next
                        Catch ex As Exception
                            ReportError("LoadForm caused an error: " & ex.Message, ex)
                        End Try
                    End Sub)

            End SyncLock

            ' Ensure Keyboard Module is loaded, 
            ' to create a global hanler for the PreviewKeyDown event
            Dim __ = Keyboard.LastKey

            Return form_Name
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

        Private Shared Function GetControlText(control As UIElement) As String
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

        ''' <summary>
        ''' Creats a new form with the given name, and adds it to the forms collection.
        ''' </summary>
        ''' <param name="formName">the name of the form</param>
        ''' <returns>the name of the form</returns>
        Public Shared Function AddForm(formName As Primitive) As Primitive
            Try
                Dim frm = GetForm(formName)
                If frm Is Nothing Then
                    Dim xaml As String = $"<Canvas Name=""{formName}"" Width=""700"" Height=""500"" xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""/>"
                    LoadForm(formName, xaml)
                End If

            Catch ex As Exception
                ReportError("AddForm cuased this error: " & ex.Message, ex)
            End Try
            Return formName
        End Function

        Private Shared Sub Form_Closed(sender As Object, e As EventArgs)
            Dim win = CType(sender, Window)
            Dim formName = win.Name
            _forms.Remove(formName)
        End Sub

        Private Shared Sub Form_Closing(sender As Object, e As CancelEventArgs)
            Dim win = CType(sender, Window)
            Dim formName = win.Name

            Dim CodeFilePath = Environment.GetCommandLineArgs(0)
            Dim docName = IO.Path.GetFileNameWithoutExtension(CodeFilePath) & ".xaml"
            CodeFilePath = IO.Path.GetDirectoryName(CodeFilePath)
            Dim newXamlPath = IO.Path.Combine(CodeFilePath, docName)

            Try
                ' If the form is created from code, save its design to .xml file
                If Not IO.File.Exists(newXamlPath) Then
                    Dim canvas = win.Content
                    IO.File.WriteAllText(newXamlPath, XamlWriter.Save(canvas))
                End If
            Finally
            End Try

        End Sub

        Private Shared Function LoadContent(xamlPath As String) As Canvas
            If IO.Path.GetPathRoot(xamlPath) = "" Then
                If AppPath.ToString() <> "" Then
                    xamlPath = IO.Path.Combine(AppPath, xamlPath)
                Else
                    Dim d = Environment.GetCommandLineArgs(0)
                    d = IO.Path.GetDirectoryName(d)
                    Dim xamlPath2 = IO.Path.Combine(d, xamlPath)
                    If IO.File.Exists(xamlPath2) Then xamlPath = xamlPath2
                End If
            End If

            Dim xaml = IO.File.ReadAllText(xamlPath, System.Text.Encoding.UTF8)
            If xaml.Contains("wpf/xaml/WpfDialogs") Then
                xaml = "<Canvas " & "xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""c""" & xaml.Substring(7)
            End If
            Dim stream = New IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(xaml))

            Dim canvas As Canvas = Nothing

            canvas = CType(XamlReader.Load(stream), Canvas)

            Return canvas
        End Function

        ''' <summary>
        ''' Shows a message box dialog.
        ''' </summary>
        ''' <param name="message">the text to dislpay on the message box</param>
        ''' <param name="title">the title to display of the dialog box</param>
        Public Shared Sub ShowMessage(message As Primitive, title As Primitive)
            Try
                SmallBasicApplication.Invoke(Sub() System.Windows.MessageBox.Show(message.ToString(), title.ToString()))
            Catch ex As Exception
                ReportError("ShowMessage caused this error: " & ex.Message, ex)
            End Try
        End Sub
    End Class
End Namespace