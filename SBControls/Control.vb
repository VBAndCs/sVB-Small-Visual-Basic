Imports Wpf = System.Windows.Controls
Imports Microsoft.SmallBasic.Library

<SmallBasicType>
Public Module Control

    Public Function GetLeft(formName As Primitive, controlName As Primitive) As Primitive
        Dispatcher.Invoke(Sub() GetLeft = Canvas.GetLeft(GetControl(formName, controlName)))
    End Function

    Public Sub SetLeft(formName As Primitive, controlName As Primitive, value As Primitive)
        Dispatcher.Invoke(Sub() Canvas.SetLeft(GetControl(formName, controlName), value))
    End Sub


    Public Function GetTop(formName As Primitive, controlName As Primitive) As Primitive
        Dispatcher.Invoke(Sub() GetTop = Canvas.GetTop(GetControl(formName, controlName)))
    End Function

    Public Sub SetTop(formName As Primitive, controlName As Primitive, value As Primitive)
        Dispatcher.Invoke(Sub() Canvas.SetTop(GetControl(formName, controlName), value))
    End Sub

    Public Function GetWidth(formName As Primitive, controlName As Primitive) As Primitive
        Dispatcher.Invoke(Sub() GetWidth = GetControl(formName, controlName).ActualWidth)
    End Function

    Public Sub SetWidth(formName As Primitive, controlName As Primitive, value As Primitive)
        Dispatcher.Invoke(Sub() GetControl(formName, controlName).Width = value)
    End Sub

    Public Function GetHeight(formName As Primitive, controlName As Primitive) As Primitive
        Dispatcher.Invoke(Sub() GetHeight = GetControl(formName, controlName).ActualHeight)
    End Function

    Public Sub SetHeight(formName As Primitive, controlName As Primitive, value As Primitive)
        Dispatcher.Invoke(Sub() Dispatcher.Invoke(Sub() GetControl(formName, controlName).Height = value))
    End Sub
    Public Function GetEnabled(formName As Primitive, controlName As Primitive) As Primitive
        Dispatcher.Invoke(Sub() GetEnabled = GetControl(formName, controlName).IsEnabled)
    End Function

    Public Sub SetEnabled(formName As Primitive, controlName As Primitive, value As Primitive)
        Dispatcher.Invoke(Sub() GetControl(formName, controlName).IsEnabled = value)
    End Sub

    Public Function GetVisible(formName As Primitive, controlName As Primitive) As Primitive
        Dispatcher.Invoke(Sub() GetVisible = GetControl(formName, controlName).IsVisible)
    End Function

    Public Sub SetVisible(formName As Primitive, controlName As Primitive, value As Primitive)
        Dispatcher.Invoke(Sub() GetControl(formName, controlName).Visibility = If(value, Visibility.Visible, Visibility.Hidden))
    End Sub

    Public Function GetBackColor(formName As Primitive, controlName As Primitive) As Primitive
        Dim brush = GetControl(formName, controlName).Background

    End Function

    Public Sub SetBackColor(formName As Primitive, controlName As Primitive, value As Primitive)
        ' c = convert value to color
        ' GetControl(formName, controlName).Background = New SolidColorBrush(c)
    End Sub


#Region "Events"
    Public ReadOnly Property SenderForm As Primitive
    Public ReadOnly Property SenderControl As Primitive


    Public Sub HandleEvents(FormName As Primitive, ControlName As Primitive)
        _SenderForm = FormName
        _SenderControl = ControlName
    End Sub

    Sub EventsHandler(sender As Wpf.Control, handler As SmallBasicCallback)
        _SenderControl = sender.Name
        If TypeOf sender Is Window Then
            _SenderForm = sender.Name
        Else
            _SenderForm = CType(GetParent(sender, GetType(Window)), Window).Name
        End If
        handler()
    End Sub

    Public Custom Event OnMouseLeftDown As SmallBasicCallback
        AddHandler(handler As SmallBasicCallback)
            Dim VisualElement = GetControl(_SenderForm, _SenderControl)
            AddHandler VisualElement.PreviewMouseLeftButtonDown,
                Sub(Sender As Wpf.Control, e As EventArgs) EventsHandler(Sender, handler)
        End AddHandler

        RemoveHandler(handler As SmallBasicCallback)
        End RemoveHandler

        RaiseEvent()
        End RaiseEvent
    End Event

    Public Custom Event OnClick As SmallBasicCallback
        AddHandler(handler As SmallBasicCallback)
            Dim VisualElement = GetControl(_SenderForm, _SenderControl)
            AddHandler VisualElement.MouseLeftButtonUp,
                Sub(Sender As Wpf.Control, e As EventArgs) EventsHandler(Sender, handler)
        End AddHandler

        RemoveHandler(handler As SmallBasicCallback)
        End RemoveHandler

        RaiseEvent()
        End RaiseEvent
    End Event

    Public Custom Event OnMouseLeftUp As SmallBasicCallback
        AddHandler(handler As SmallBasicCallback)
            Dim VisualElement = GetControl(_SenderForm, _SenderControl)
            AddHandler VisualElement.PreviewMouseLeftButtonUp,
                Sub(Sender As Wpf.Control, e As EventArgs) EventsHandler(Sender, handler)
        End AddHandler

        RemoveHandler(handler As SmallBasicCallback)
        End RemoveHandler

        RaiseEvent()
        End RaiseEvent
    End Event

    Public Custom Event OnDoubleClick As SmallBasicCallback
        AddHandler(handler As SmallBasicCallback)
            Dim VisualElement = GetControl(_SenderForm, _SenderControl)
            AddHandler VisualElement.MouseDoubleClick,
                Sub(Sender As Wpf.Control, e As EventArgs) EventsHandler(Sender, handler)
        End AddHandler

        RemoveHandler(handler As SmallBasicCallback)
        End RemoveHandler

        RaiseEvent()
        End RaiseEvent
    End Event

    Public Custom Event OnMouseRightDown As SmallBasicCallback
        AddHandler(handler As SmallBasicCallback)
            Dim VisualElement = GetControl(_SenderForm, _SenderControl)
            AddHandler VisualElement.PreviewMouseRightButtonDown,
                Sub(Sender As Wpf.Control, e As EventArgs) EventsHandler(Sender, handler)
        End AddHandler

        RemoveHandler(handler As SmallBasicCallback)
        End RemoveHandler

        RaiseEvent()
        End RaiseEvent
    End Event

    Public Custom Event OnMouseRightUp As SmallBasicCallback
        AddHandler(handler As SmallBasicCallback)
            Dim VisualElement = GetControl(_SenderForm, _SenderControl)
            AddHandler VisualElement.PreviewMouseRightButtonUp,
                Sub(Sender As Wpf.Control, e As EventArgs) EventsHandler(Sender, handler)
        End AddHandler

        RemoveHandler(handler As SmallBasicCallback)
        End RemoveHandler

        RaiseEvent()
        End RaiseEvent
    End Event

    Public Custom Event OnMouseMove As SmallBasicCallback
        AddHandler(handler As SmallBasicCallback)
            Dim VisualElement = GetControl(_SenderForm, _SenderControl)
            AddHandler VisualElement.PreviewMouseMove,
                Sub(Sender As Wpf.Control, e As EventArgs) EventsHandler(Sender, handler)
        End AddHandler

        RemoveHandler(handler As SmallBasicCallback)
        End RemoveHandler

        RaiseEvent()
        End RaiseEvent
    End Event

    Public Custom Event OnMouseWheel As SmallBasicCallback
        AddHandler(handler As SmallBasicCallback)
            Dim VisualElement = GetControl(_SenderForm, _SenderControl)
            AddHandler VisualElement.PreviewMouseWheel,
                Sub(Sender As Wpf.Control, e As EventArgs) EventsHandler(Sender, handler)
        End AddHandler

        RemoveHandler(handler As SmallBasicCallback)
        End RemoveHandler

        RaiseEvent()
        End RaiseEvent
    End Event

    Public Custom Event OnMouseEnter As SmallBasicCallback
        AddHandler(handler As SmallBasicCallback)
            Dim VisualElement = GetControl(_SenderForm, _SenderControl)
            AddHandler VisualElement.MouseEnter,
                Sub(Sender As Wpf.Control, e As EventArgs) EventsHandler(Sender, handler)
        End AddHandler

        RemoveHandler(handler As SmallBasicCallback)
        End RemoveHandler

        RaiseEvent()
        End RaiseEvent
    End Event

    Public Custom Event OnMouseLeave As SmallBasicCallback
        AddHandler(handler As SmallBasicCallback)
            Dim VisualElement = GetControl(_SenderForm, _SenderControl)
            AddHandler VisualElement.MouseLeave,
                Sub(Sender As Wpf.Control, e As EventArgs) EventsHandler(Sender, handler)
        End AddHandler

        RemoveHandler(handler As SmallBasicCallback)
        End RemoveHandler

        RaiseEvent()
        End RaiseEvent
    End Event

    Public Custom Event OnKeyDown As SmallBasicCallback
        AddHandler(handler As SmallBasicCallback)
            Dim VisualElement = GetControl(_SenderForm, _SenderControl)
            AddHandler VisualElement.PreviewKeyDown,
                Sub(Sender As Wpf.Control, e As EventArgs) EventsHandler(Sender, handler)
        End AddHandler

        RemoveHandler(handler As SmallBasicCallback)
        End RemoveHandler

        RaiseEvent()
        End RaiseEvent
    End Event

    Public Custom Event OnKeyUp As SmallBasicCallback
        AddHandler(handler As SmallBasicCallback)
            Dim VisualElement = GetControl(_SenderForm, _SenderControl)
            AddHandler VisualElement.PreviewKeyUp,
                Sub(Sender As Wpf.Control, e As EventArgs) EventsHandler(Sender, handler)
        End AddHandler

        RemoveHandler(handler As SmallBasicCallback)
        End RemoveHandler

        RaiseEvent()
        End RaiseEvent
    End Event


#End Region


    Friend Function GetControl(formName As String, controlName As String) As Wpf.Control
        If Not _forms.ContainsKey(formName) Then
            Throw New ArgumentException($"There is no form named `{formName}`.")
        End If

        Dim _controls = _forms(formName)
        If controlName = "" Then Return _controls(formName)

        If Not _controls.ContainsKey(controlName) Then
            Throw New ArgumentException($"There is no control named `{controlName}` on form 'formName'.")
        End If
        Return _controls(controlName)
    End Function

    Private Function GetParent(child As DependencyObject, parentType As Type) As UIElement
        If child Is Nothing Then Return Nothing
        Dim p = child
        Do
            p = VisualTreeHelper.GetParent(p)
            If p Is Nothing Then Return Nothing
            If p.GetType Is parentType Then Return p
        Loop
    End Function
End Module


