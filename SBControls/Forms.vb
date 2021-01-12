Imports System.Windows.Markup
Imports System.Windows.Threading
Imports Microsoft.SmallBasic.Library
Imports Microsoft.SmallBasic.Library.Internal
Imports Wpf = System.Windows.Controls
Imports ControlsDictionay = System.Collections.Generic.Dictionary(Of String, System.Windows.Controls.Control)
Imports DiagramHelper
Imports DiagramHelper.Designer

<SmallBasicType>
Public Module Forms
    <System.Runtime.CompilerServices.Extension()>
    Public Iterator Function GetChildren(ByVal parent As UIElement, ByVal Optional recurse As Boolean = True) As IEnumerable(Of UIElement)
        If parent IsNot Nothing Then
            Dim count As Integer = VisualTreeHelper.GetChildrenCount(parent)

            For i As Integer = 0 To count - 1
                Dim child = TryCast(VisualTreeHelper.GetChild(parent, i), UIElement)

                If child IsNot Nothing Then
                    Yield child

                    If recurse Then
                        For Each grandChild In child.GetChildren(True)
                            Yield grandChild
                        Next
                    End If
                End If
            Next
        End If
    End Function

    Friend _forms As New Dictionary(Of String, ControlsDictionay)

    Function GetForm(name As String) As System.Windows.Window
        If Not _forms.ContainsKey(name) Then
            Throw New ArgumentException($"There is no form named `{name}`.")
        End If

        Dim wnd = _forms(name)(name)
        If wnd Is Nothing Then
            Throw New ArgumentException($"`{name}` is not a form or it is closed")
        End If

        Return wnd
    End Function

    Public Function GetForms() As Primitive
        Dim map = New Dictionary(Of Primitive, Primitive)
        Dim num = 1
        For Each key In _forms.Keys
            map(num) = key
            num += 1
        Next
        Return Primitive.ConvertFromMap(map)
    End Function

    Public Function LoadForm(name As Primitive, xamlPath As Primitive) As Primitive
        If CStr(name) = "" Then
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
                   .Content = LoadContent(xamlPath)
               }

                Dim _controls = New ControlsDictionay()
                _controls.Add(name, wnd)
                _forms.Add(name, _controls)

                ' Add control names:
                For Each ui In CType(wnd.Content, UIElement).GetChildren()
                    Dim c = TryCast(ui, Wpf.Control)
                    If c IsNot Nothing AndAlso c.Name <> "" Then
                        _controls.Add(c.Name, c)
                    End If
                Next
            End Sub)

        Return name
    End Function

    Public Function AddForm(name As Primitive) As Primitive
        If CStr(name) = "" Then
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


    Private Function LoadContent(FileName As String) As UIElement
        Dim stream = IO.File.Open(FileName, IO.FileMode.Open)
        Dim reader As New XamlReader()
        Return CType(reader.LoadAsync(stream), UIElement)
    End Function


    Dim _dispatcher As Dispatcher
    Friend ReadOnly Property Dispatcher() As Dispatcher
        Get
            If _dispatcher Is Nothing Then
                Dim prop = GetType(SmallBasicApplication).GetProperty("Dispatcher", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Static)
                _dispatcher = prop.GetValue(Nothing)
            End If

            Return _dispatcher
        End Get
    End Property


End Module