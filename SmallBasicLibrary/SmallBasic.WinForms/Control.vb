Imports Wpf = System.Windows.Controls
Imports Microsoft.SmallBasic.Library
Imports System.Windows
Imports System.Windows.Media

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class Control

        ''' <summary>
        ''' The x-pos of the control on its parent control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetLeft(formName As Primitive, controlName As Primitive) As Primitive
            Forms.Dispatcher.Invoke(Sub() GetLeft = Wpf.Canvas.GetLeft(GetControl(formName, controlName)))
        End Function

        <ExProperty>
        Public Shared Sub SetLeft(formName As Primitive, controlName As Primitive, value As Primitive)
            Forms.Dispatcher.Invoke(Sub() Wpf.Canvas.SetLeft(GetControl(formName, controlName), value))
        End Sub

        ''' <summary>
        ''' The y-pos of the control on its parent control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetTop(formName As Primitive, controlName As Primitive) As Primitive
            Forms.Dispatcher.Invoke(Sub() GetTop = Wpf.Canvas.GetTop(GetControl(formName, controlName)))
        End Function

        <ExProperty>
        Public Shared Sub SetTop(formName As Primitive, controlName As Primitive, value As Primitive)
            Forms.Dispatcher.Invoke(Sub() Wpf.Canvas.SetTop(GetControl(formName, controlName), value))
        End Sub

        ''' <summary>
        ''' The width of the control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetWidth(formName As Primitive, controlName As Primitive) As Primitive
            Forms.Dispatcher.Invoke(Sub() GetWidth = GetControl(formName, controlName).ActualWidth)
        End Function

        <ExProperty>
        Public Shared Sub SetWidth(formName As Primitive, controlName As Primitive, value As Primitive)
            Forms.Dispatcher.Invoke(Sub() GetControl(formName, controlName).Width = value)
        End Sub

        ''' <summary>
        ''' The height of the control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetHeight(formName As Primitive, controlName As Primitive) As Primitive
            Forms.Dispatcher.Invoke(Sub() GetHeight = GetControl(formName, controlName).ActualHeight)
        End Function

        <ExProperty>
        Public Shared Sub SetHeight(formName As Primitive, controlName As Primitive, value As Primitive)
            Forms.Dispatcher.Invoke(Sub() Forms.Dispatcher.Invoke(Sub() GetControl(formName, controlName).Height = value))
        End Sub

        ''' <summary>
        ''' When its value = 1 (or True), user can interact with the control.
        ''' When its value = 0 (or False),  the control is disabled, and user can't interact with it.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetEnabled(formName As Primitive, controlName As Primitive) As Primitive
            Forms.Dispatcher.Invoke(Sub() GetEnabled = GetControl(formName, controlName).IsEnabled)
        End Function

        <ExProperty>
        Public Shared Sub SetEnabled(formName As Primitive, controlName As Primitive, value As Primitive)
            Forms.Dispatcher.Invoke(Sub() GetControl(formName, controlName).IsEnabled = value)
        End Sub

        ''' <summary>
        ''' When its value = 1 (or True), the control is shown at the form.
        ''' When its value = 0 (or False),  the control is hidden.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetVisible(formName As Primitive, controlName As Primitive) As Primitive
            Forms.Dispatcher.Invoke(Sub() GetVisible = GetControl(formName, controlName).IsVisible)
        End Function

        <ExProperty>
        Public Shared Sub SetVisible(formName As Primitive, controlName As Primitive, value As Primitive)
            Forms.Dispatcher.Invoke(Sub() GetControl(formName, controlName).Visibility = If(value, Visibility.Visible, Visibility.Hidden))
        End Sub


        Private Shared ReadOnly BackColorProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("BackColor",
                           GetType(String), GetType(Wpf.Control))

        ''' <summary>
        ''' The backgeound color of the control.
        ''' Use values from the Color object, such as Color.Yellow
        ''' </summary>
        <ExProperty>
        Public Shared Function GetBackColor(formName As Primitive, controlName As Primitive) As Primitive
            Forms.Dispatcher.Invoke(
           Sub()
               Dim c = GetControl(formName, controlName)
               Dim brush = TryCast(c.Background, SolidColorBrush)
               If brush IsNot Nothing Then
                   GetBackColor = brush.Color.ToString()
               Else
                   GetBackColor = c.GetValue(BackColorProperty)
               End If
           End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetBackColor(formName As Primitive, controlName As Primitive, value As Primitive)
            Forms.Dispatcher.Invoke(
           Sub()
               Dim c = GetControl(formName, controlName)
               Dim _color = Color.FromString(value)
               c.Background = New SolidColorBrush(_color)
               c.SetValue(BackColorProperty, value.ToString())
           End Sub)
        End Sub


        Private Shared ReadOnly ForeColorProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("ForeColor",
                           GetType(String), GetType(Wpf.Control))

        ''' <summary>
        ''' The foregeound color used to draw the text of the control.
        ''' Use values from the Color object, such as Color.Red
        ''' </summary>
        <ExProperty>
        Public Shared Function GetForeColor(formName As Primitive, controlName As Primitive) As Primitive
            Forms.Dispatcher.Invoke(
           Sub()
               Dim c = GetControl(formName, controlName)
               Dim brush = TryCast(c.Foreground, SolidColorBrush)
               If brush IsNot Nothing Then
                   GetForeColor = brush.Color.ToString()
               Else
                   GetForeColor = c.GetValue(ForeColorProperty)
               End If
           End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetForeColor(formName As Primitive, controlName As Primitive, value As Primitive)
            Forms.Dispatcher.Invoke(
           Sub()
               Dim c = GetControl(formName, controlName)
               Dim _color = Color.FromString(value)
               c.Foreground = New SolidColorBrush(_color)
               c.SetValue(ForeColorProperty, value.ToString())
           End Sub)
        End Sub

        ''' <summary>
        ''' The mouse x-pos relative to the control. When mouse is over the control, this value lies between 0 and the control's width.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetMouseX(formName As Primitive, controlName As Primitive) As Primitive
            Forms.Dispatcher.Invoke(
            Sub()
                Dim c = GetControl(formName, controlName)
                GetMouseX = System.Math.Round(Input.Mouse.GetPosition(c).X)
            End Sub)
        End Function

        ''' <summary>
        ''' The mouse y-pos relative to the control. When mouse is over the control, this value lies between 0 and the control's.height.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetMouseY(formName As Primitive, controlName As Primitive) As Primitive
            Forms.Dispatcher.Invoke(
            Sub()
                Dim c = GetControl(formName, controlName)
                GetMouseY = System.Math.Round(Input.Mouse.GetPosition(c).Y)
            End Sub)
        End Function

#Region "Events"
        Public Shared ReadOnly Property SenderForm As Primitive
        Public Shared ReadOnly Property SenderControl As Primitive

        <ExMethod>
        Public Shared Sub HandleEvents(FormName As Primitive, ControlName As Primitive)
            _SenderForm = FormName
            _SenderControl = ControlName
        End Sub


        Shared Sub EventsHandler(sender As Wpf.Control, handler As SmallBasicCallback)
            _SenderControl = sender.Name
            If TypeOf sender Is Window Then
                _SenderForm = sender.Name
            Else
                _SenderForm = CType(GetParent(sender, GetType(Window)), Window).Name
            End If
            handler()
        End Sub

        ''' <summary>
        ''' Fired when user presses the left mouse-button down
        ''' </summary>
        Public Shared Custom Event OnMouseLeftDown As SmallBasicCallback
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

        ''' <summary>
        ''' Fired when user releases the left mouse-button
        ''' </summary>
        Public Shared Custom Event OnClick As SmallBasicCallback
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

        ''' <summary>
        '''     ''' Fired when user releases the left mouse-button
        ''' </summary>
        Public Shared Custom Event OnMouseLeftUp As SmallBasicCallback
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

        ''' <summary>
        '''     ''' Fired when user double-clicks the mouse-button
        ''' </summary>
        Public Shared Custom Event OnDoubleClick As SmallBasicCallback
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

        ''' <summary>
        ''' Fired when user presses the right mouse-button down
        ''' </summary>
        Public Shared Custom Event OnMouseRightDown As SmallBasicCallback
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

        ''' <summary>
        ''' Fired when user releases the right mouse-button
        ''' </summary>
        Public Shared Custom Event OnMouseRightUp As SmallBasicCallback
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

        ''' <summary>
        ''' Fired when the mouse pointer moves over the control.
        ''' </summary>
        Public Shared Custom Event OnMouseMove As SmallBasicCallback
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

        ''' <summary>
        ''' Fired when user moves the mouse wheel
        ''' </summary>
        Public Shared Custom Event OnMouseWheel As SmallBasicCallback
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

        ''' <summary>
        ''' Fired when the mouse pointer enters the control area.
        ''' </summary>
        Public Shared Custom Event OnMouseEnter As SmallBasicCallback
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

        ''' <summary>
        ''' Fired when the mouse pointer leaves the control area.
        ''' </summary>
        Public Shared Custom Event OnMouseLeave As SmallBasicCallback
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

        ''' <summary>
        ''' Fired when the user presses a keyboard-ky down
        ''' </summary>
        Public Shared Custom Event OnKeyDown As SmallBasicCallback
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

        ''' <summary>
        ''' Fired when the user releases a keyboard-ky.
        ''' </summary>
        Public Shared Custom Event OnKeyUp As SmallBasicCallback
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


        Friend Shared Function GetControl(formName As String, controlName As String) As Wpf.Control
            formName = formName.ToLower()
            If Not Forms._forms.ContainsKey(formName) Then
                Throw New ArgumentException($"There is no form named `{formName}`.")
            End If


            Dim _controls = Forms._forms(formName)
            If controlName = "" Then Return _controls(formName)

            Dim name = controlName.ToLower()
            If Not _controls.ContainsKey(name) Then
                Throw New ArgumentException($"There is no control named `{controlName}` on form {formName}.")
            End If
            Return _controls(name)
        End Function

        Private Shared Function GetParent(child As DependencyObject, parentType As Type) As UIElement
            If child Is Nothing Then Return Nothing
            Dim p = child
            Do
                p = VisualTreeHelper.GetParent(p)
                If p Is Nothing Then Return Nothing
                If p.GetType Is parentType Then Return p
            Loop
        End Function
    End Class


End Namespace