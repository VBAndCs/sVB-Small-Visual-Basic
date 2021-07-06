Imports System.Windows.Markup
Imports System.Windows.Threading
Imports Microsoft.SmallBasic.Library
Imports Microsoft.SmallBasic.Library.Internal
Imports Wpf = System.Windows.Controls
Imports ControlsDictionay = System.Collections.Generic.Dictionary(Of String, System.Windows.Controls.Control)


<SmallBasicType>
Public NotInheritable Class Forms

    Friend Shared _forms As New Dictionary(Of String, ControlsDictionay)

    Shared Function GetForm(name As String) As Window
        name = name.ToLower()
        If Not _forms.ContainsKey(name) Then
            Throw New ArgumentException($"There is no form named `{name}`.")
        End If

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

    Public Shared Function LoadForm(formName As Primitive, xamlPath As Primitive) As Primitive
        Dim name = CStr(formName).ToLower()
        If name = "" Then
            Throw New ArgumentException("Form name can't be an empty string.")
        End If

        If _forms.ContainsKey(name) Then
            Throw New ArgumentException($"There is already a form named `{name}`.")
        End If

        If Not IO.File.Exists(xamlPath) Then
            If IO.Path.GetPathRoot(xamlPath) = "" Then
                Dim d = AppDomain.CurrentDomain.BaseDirectory
                Dim xamlPath2 = IO.Path.Combine(d, xamlPath)
                If IO.File.Exists(xamlPath2) Then
                    xamlPath = xamlPath2
                ElseIf AppPath.ToString() <> "" Then
                    xamlPath = IO.Path.Combine(AppPath, xamlPath)
                End If
            End If
        End If

        Dispatcher.Invoke(
            Sub()
                Dim wnd As New Window() With {
                   .Width = 500,
                   .Height = 300,
                   .Name = name,
                   .Content = LoadContent(xamlPath)
               }

                Dim _controls = New ControlsDictionay()
                _controls.Add(name, wnd)
                _forms.Add(name, _controls)

                ' Add control names:
                For Each ui In CType(wnd.Content, UIElement).GetChildren()
                    Dim c = TryCast(ui, Wpf.Control)
                    If c IsNot Nothing AndAlso c.Name <> "" Then
                        _controls.Add(c.Name.ToLower(), c)
                    End If
                Next
            End Sub)

        ' Ensure Keyboard Module is loaded, 
        ' to create a global hanler for the PreviewKeyDown event
        Dim __ = Keyboard.LastKey

        Return name
    End Function

    Public Shared Function AddForm(formName As Primitive) As Primitive
        Dim name = CStr(formName).ToLower()
        If name = "" Then
            Throw New ArgumentException("Form name can't be an empty string.")
        End If

        If _forms.ContainsKey(name) Then
            Throw New ArgumentException($"There is already a form named `{name}`.")
        End If

        Dispatcher.Invoke(
            Sub()
                Dim wnd As New Window() With {
                   .Width = 500,
                   .Height = 300,
                   .Name = name,
                   .Content = New Canvas()
                }

                Dim _controls = New ControlsDictionay()
                _controls.Add(name, wnd)
                _forms.Add(name, _controls)
            End Sub)

        Return name
    End Function

    Private Shared Function LoadContent(FileName As String) As UIElement
        Dim stream = IO.File.Open(FileName, IO.FileMode.Open)
        Dim reader As New XamlReader()
        Return CType(reader.LoadAsync(stream), UIElement)
    End Function


    Private Shared _dispatcher As Dispatcher
    Friend Shared ReadOnly Property Dispatcher() As Dispatcher
        Get
            If _dispatcher Is Nothing Then
                Dim prop = GetType(SmallBasicApplication).GetProperty("Dispatcher", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Static)
                _dispatcher = prop.GetValue(Nothing)
            End If

            Return _dispatcher
        End Get
    End Property

    Public Shared Sub ShowMessage(message As Primitive, title As Primitive)
        MessageBox.Show(message.ToString(), title.ToString())
    End Sub
End Class