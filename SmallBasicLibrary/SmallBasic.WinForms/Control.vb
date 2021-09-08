Option Explicit On
Option Strict On


Imports Wpf = System.Windows.Controls
Imports Microsoft.SmallBasic.Library
Imports System.Windows
Imports System.Windows.Media
Imports App = Microsoft.SmallBasic.Library.Internal.SmallBasicApplication

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class Control
        Shared Sub ShowSubError(formName As String, controlName As String, memberName As String, msg As String)
            If controlName.ToLower() = formName.ToLower Then
                MsgBox($"Calling {formName}.{memberName} caused an error: {vbCrLf}{msg}")
            Else
                MsgBox($"Calling {formName}.{controlName}.{memberName} caused an error: {vbCrLf}{msg}")
            End If
        End Sub

        Shared Sub ShowErrorMesssage(formName As String, controlName As String, memberName As String, msg As String)
            If controlName.ToLower() = formName.ToLower Then
                MsgBox($"Reading {formName}.{memberName} caused an error: {vbCrLf}{msg}")
            Else
                MsgBox($"Reading {formName}.{controlName}.{memberName} caused an error: {vbCrLf}{msg}")
            End If
        End Sub

        Shared Sub ShowErrorMesssage(formName As String, controlName As String, memberName As String, value As String, msg As String)
            If controlName.ToLower() = formName.ToLower Then
                MsgBox($"Sending `{value}` to {formName}.{memberName} caused an error: {vbCrLf}{msg}")
            Else
                MsgBox($"Sending `{value}` to {formName}.{controlName}.{memberName} caused an error: {vbCrLf}{msg}")
            End If
        End Sub


        ''' <summary>
        ''' Gets the name of the control.           
        ''' </summary>
        ''' <remarks>You can't change the control name at runtime. Use the designer to change the name</remarks>
        <ExProperty>
        Public Shared Function GetName(formName As Primitive, controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetName = GetControl(formName, controlName).Name
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Name", ex.Message)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' The x-pos of the control on its parent control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetLeft(formName As Primitive, controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetLeft = Wpf.Canvas.GetLeft(GetControl(formName, controlName))
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Left", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetLeft(formName As Primitive, controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(formName, controlName)
                        ' Remove any animation effect to allow setting the new value
                        obj.BeginAnimation(Wpf.Canvas.LeftProperty, Nothing)
                        Wpf.Canvas.SetLeft(obj, value)
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Left", value, ex.Message)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The y-pos of the control on its parent control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetTop(formName As Primitive, controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetTop = Wpf.Canvas.GetTop(GetControl(formName, controlName))
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Top", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetTop(formName As Primitive, controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(formName, controlName)
                        ' Remove any animation effect to allow setting the new value
                        obj.BeginAnimation(Wpf.Canvas.TopProperty, Nothing)
                        Wpf.Canvas.SetTop(obj, value)
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Top", value, ex.Message)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The width of the control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetWidth(formName As Primitive, controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(formName, controlName)
                        If TypeOf obj Is Window Then
                            GetWidth = CType(CType(obj, Window).Content, Wpf.Canvas).ActualWidth
                        Else
                            GetWidth = obj.ActualWidth
                        End If
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Width", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetWidth(formName As Primitive, controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(formName, controlName)
                        If TypeOf obj Is Window Then
                            ' Remove any animation effect to allow setting the new value
                            obj.BeginAnimation(Wpf.Canvas.WidthProperty, Nothing)
                            CType(CType(obj, Window).Content, Wpf.Canvas).Width = value
                        Else
                            ' Remove any animation effect to allow setting the new value
                            obj.BeginAnimation(Wpf.Control.WidthProperty, Nothing)
                            obj.Width = value
                        End If
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Width", value, ex.Message)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The height of the control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetHeight(formName As Primitive, controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(formName, controlName)
                        If TypeOf obj Is Window Then
                            GetHeight = CType(CType(obj, Window).Content, Wpf.Canvas).ActualHeight
                        Else
                            GetHeight = obj.ActualHeight
                        End If
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Height", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetHeight(formName As Primitive, controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(formName, controlName)
                        If TypeOf obj Is Window Then
                            ' Remove any animation effect to allow setting the new value
                            obj.BeginAnimation(Wpf.Canvas.HeightProperty, Nothing)
                            CType(CType(obj, Window).Content, Wpf.Canvas).Height = value
                        Else
                            ' Remove any animation effect to allow setting the new value
                            obj.BeginAnimation(Wpf.Control.HeightProperty, Nothing)
                            obj.Height = value
                        End If
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Height", value, ex.Message)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' When its value = 1 (or True), user can interact with the control.
        ''' When its value = 0 (or False),  the control is disabled, and user can't interact with it.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetEnabled(formName As Primitive, controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetEnabled = GetControl(formName, controlName).IsEnabled
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Ebabled", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetEnabled(formName As Primitive, controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetControl(formName, controlName).IsEnabled = value
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Enabled", value, ex.Message)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' When its value = 1 (or True), the control is shown at the form.
        ''' When its value = 0 (or False),  the control is hidden.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetVisible(formName As Primitive, controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetVisible = GetControl(formName, controlName).IsVisible
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Visible", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetVisible(formName As Primitive, controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetControl(formName, controlName).Visibility = If(value, Visibility.Visible, Visibility.Hidden)
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Visible", value, ex.Message)
                    End Try
                End Sub)
        End Sub


        Private Shared ReadOnly BackColorProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("BackColor",
                           GetType(Media.Color), GetType(Wpf.Control),
                           New PropertyMetadata(SystemColors.ControlColor, AddressOf BackColor_Changed))

        Private Shared Sub BackColor_Changed(c As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim _color = CType(e.NewValue, Media.Color)
            If TypeOf c Is Window Then
                CType(CType(c, Window).Content, Wpf.Canvas).Background = New SolidColorBrush(_color)
            End If

            CType(c, Wpf.Control).Background = New SolidColorBrush(_color)
        End Sub

        ''' <summary>
        ''' The backgeound color of the control.
        ''' Use values from the Color object, such as Color.Yellow
        ''' </summary>
        <ExProperty>
        Public Shared Function GetBackColor(formName As Primitive, controlName As Primitive) As Primitive
            App.Invoke(
                 Sub()
                     Try
                         Dim c = GetControl(formName, controlName)
                         GetBackColor = GetBackColor(c).ToString()
                     Catch ex As Exception
                         ShowErrorMesssage(formName, controlName, "BackColor", ex.Message)
                     End Try
                 End Sub)
        End Function

        Public Shared Function GetBackColor(control As Wpf.Control) As Media.Color
            Dim brush As SolidColorBrush
            If TypeOf control Is Window Then
                Dim canvas = CType(CType(control, Window).Content, Wpf.Canvas)
                brush = TryCast(canvas.Background, SolidColorBrush)
            Else
                brush = TryCast(control.Background, SolidColorBrush)
            End If

            If Brush IsNot Nothing Then
                Return brush.Color
            Else
                Return CType(control.GetValue(BackColorProperty), Media.Color)
            End If
        End Function

        <ExProperty>
        Public Shared Sub SetBackColor(formName As Primitive, controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim _color = Color.FromString(value)
                        Dim c = GetControl(formName, controlName)
                        c.SetValue(BackColorProperty, _color)
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "BackColor", value, ex.Message)
                    End Try
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
            App.Invoke(
                  Sub()
                      Try
                          Dim c = GetControl(formName, controlName)
                          Dim brush = TryCast(c.Foreground, SolidColorBrush)
                          If brush IsNot Nothing Then
                              GetForeColor = brush.Color.ToString()
                          Else
                              GetForeColor = CStr(c.GetValue(ForeColorProperty))
                          End If
                      Catch ex As Exception
                          ShowErrorMesssage(formName, controlName, "ForeColor", ex.Message)
                      End Try
                  End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetForeColor(formName As Primitive, controlName As Primitive, value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Dim c = GetControl(formName, controlName)
                           Dim _color = Color.FromString(value)
                           c.Foreground = New SolidColorBrush(_color)
                           c.SetValue(ForeColorProperty, value.ToString())
                       Catch ex As Exception
                           ShowErrorMesssage(formName, controlName, "ForeColor", value, ex.Message)
                       End Try
                   End Sub)
        End Sub

        ''' <summary>
        ''' The mouse x-pos relative to the control. When mouse is over the control, this value lies between 0 and the control's width.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetMouseX(formName As Primitive, controlName As Primitive) As Primitive
            App.Invoke(
                   Sub()
                       Try
                           Dim c = GetControl(formName, controlName)
                           GetMouseX = System.Math.Round(Input.Mouse.GetPosition(c).X)
                       Catch ex As Exception
                           ShowErrorMesssage(formName, controlName, "MouseX", ex.Message)
                       End Try
                   End Sub)
        End Function

        ''' <summary>
        ''' The mouse y-pos relative to the control. When mouse is over the control, this value lies between 0 and the control's.height.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetMouseY(formName As Primitive, controlName As Primitive) As Primitive
            App.Invoke(
                    Sub()
                        Try
                            Dim c = GetControl(formName, controlName)
                            GetMouseY = System.Math.Round(Input.Mouse.GetPosition(c).Y)
                        Catch ex As Exception
                            ShowErrorMesssage(formName, controlName, "MouseY", ex.Message)
                        End Try
                    End Sub)
        End Function


        Public Shared Function GetAngle(ByVal element As DependencyObject) As Double
            If element Is Nothing Then
                Throw New ArgumentNullException("element")
            End If

            Return CDbl(element.GetValue(AngleProperty))
        End Function


        <ExProperty>
        Public Shared Sub SetAngle(formName As Primitive, controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        SetAngle(GetControl(formName, controlName), value)
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Angle", value, ex.Message)
                    End Try
                End Sub)
        End Sub

        Public Shared Sub SetAngle(ByVal element As DependencyObject, ByVal value As Double)
            If element Is Nothing Then
                Throw New ArgumentNullException("element")
            End If

            CType(element, Wpf.Control).BeginAnimation(AngleProperty, Nothing)
            element.SetValue(AngleProperty, value)
        End Sub


        Private Shared _rotateTransformMap As New Dictionary(Of Wpf.Control, RotateTransform)

        ''' <summary>
        ''' The rotation angle of the control.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetAngle(formName As Primitive, controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetAngle = GetAngle(GetControl(formName, controlName))
                    Catch ex As Exception
                        ShowErrorMesssage(formName, controlName, "Angle", ex.Message)
                    End Try
                End Sub)
        End Function


        Public Shared ReadOnly AngleProperty As _
                               DependencyProperty = DependencyProperty.RegisterAttached("Angle",
                               GetType(Double), GetType(Wpf.Control),
                               New PropertyMetadata(0.0, AddressOf AngleChanged))

        Private Shared Sub AngleChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim obj = CType(d, Wpf.Control)
            App.BeginInvoke(
                     Sub()
                         If Not (TypeOf obj.RenderTransform Is TransformGroup) Then
                             obj.RenderTransform = New TransformGroup
                         End If

                         Dim rotation As System.Windows.Media.RotateTransform = Nothing
                         If Not _rotateTransformMap.TryGetValue(obj, rotation) Then
                             rotation = New RotateTransform
                             _rotateTransformMap(obj) = rotation
                             rotation.CenterX = obj.ActualWidth / 2.0
                             rotation.CenterY = obj.ActualHeight / 2.0
                             CType(obj.RenderTransform, TransformGroup).Children.Add(rotation)
                         End If
                         rotation.Angle = CDbl(e.NewValue)
                     End Sub)
        End Sub

        <ExMethod>
        Public Shared Sub Focus(formName As Primitive, controlName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetControl(formName, controlName).Focus()
                    Catch ex As Exception
                        ShowSubError(formName, controlName, "Focus", ex.Message)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Rotates the control to the specified angle.
        ''' </summary>
        ''' <param name="angle">
        ''' The angle to rotate the shape.
        ''' </param>
        <ExMethod>
        Public Shared Sub Rotate(formName As Primitive, controlName As Primitive, angle As Primitive)
            Try
                Dim obj = GetControl(formName, controlName)
                App.BeginInvoke(Sub() SetAngle(obj, angle))
            Catch ex As Exception
                ShowSubError(formName, controlName, "Rotate", ex.Message)
            End Try
        End Sub



#Region "Animation"

        ''' <summary>
        ''' Animates the control's backcolor to a new color.
        ''' </summary>
        ''' <param name="newColor">
        ''' The new color to animate the control's backcolor to.
        ''' </param>
        ''' <param name="duration">
        ''' The time for the animation, in milliseconds.
        ''' </param>
        <ExMethod>
        Public Shared Sub AnimateColor(formName As Primitive, controlName As Primitive, newColor As Primitive, duration As Primitive)
            Try
                Dim obj = GetControl(formName, controlName)
                App.Invoke(
                    Sub()
                        Dim animation As New Animation.ColorAnimation() With {
                            .From = GetBackColor(obj),
                            .To = Color.FromString(newColor),
                            .Duration = TimeSpan.FromMilliseconds(duration)
                        }
                        obj.BeginAnimation(BackColorProperty, animation)
                    End Sub)
            Catch ex As Exception
                ShowSubError(formName, controlName, "AnimateColor", ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Animates the control's backcolor to a new transparency.
        ''' </summary>
        ''' <param name="transparency">
        ''' The new transparency to animate the backcolor to. Use a value between 0 and 100.
        ''' </param>
        ''' <param name="duration">
        ''' The time for the animation, in milliseconds.
        ''' </param>
        <ExMethod>
        Public Shared Sub AnimateTransparency(formName As Primitive, controlName As Primitive, transparency As Primitive, duration As Primitive)
            Try
                Dim c = GetBackColor(formName, controlName)
                c = Color.SetTransparency(c, transparency)
                AnimateColor(formName, controlName, c, duration)
            Catch ex As Exception
                ShowSubError(formName, controlName, "AnimateColor", ex.Message)
            End Try
        End Sub


        ''' <summary>
        ''' Animates the control to a new position.
        ''' </summary>
        ''' <param name="x">
        ''' The x co-ordinate of the new position.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the new position.
        ''' </param>
        ''' <param name="duration">
        ''' The time for the animation, in milliseconds.
        ''' </param>
        <ExMethod>
        Public Shared Sub AnimatePos(formName As Primitive, controlName As Primitive, x As Primitive, y As Primitive, duration As Primitive)
            Try
                Dim obj = GetControl(formName, controlName)
                App.Invoke(
                    Sub()
                        GraphicsWindow.DoubleAnimateProperty(obj, Wpf.Canvas.LeftProperty, x, duration)
                        GraphicsWindow.DoubleAnimateProperty(obj, Wpf.Canvas.TopProperty, y, duration)
                    End Sub)
            Catch ex As Exception
                ShowSubError(formName, controlName, "AnimatePos", ex.Message)
            End Try
        End Sub


        ''' <summary>
        ''' Animates the control to a new size.
        ''' </summary>
        ''' <param name="width">
        ''' The new width.
        ''' </param>
        ''' <param name="height">
        ''' The new height.
        ''' </param>
        ''' <param name="duration">
        ''' The time for the animation, in milliseconds.
        ''' </param>
        <ExMethod>
        Public Shared Sub AnimateSize(formName As Primitive, controlName As Primitive, width As Primitive, height As Primitive, duration As Primitive)
            Try
                Dim obj = GetControl(formName, controlName)
                App.Invoke(
                    Sub()
                        GraphicsWindow.DoubleAnimateProperty(obj, FrameworkElement.WidthProperty, width, duration)
                        GraphicsWindow.DoubleAnimateProperty(obj, FrameworkElement.HeightProperty, height, duration)
                    End Sub)
            Catch ex As Exception
                ShowSubError(formName, controlName, "AnimateSize", ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Animates the control's rotation angle to a new angle.
        ''' </summary>
        ''' <param name="angle">
        ''' The new rotation angle.
        ''' </param>
        ''' <param name="duration">
        ''' The time for the animation, in milliseconds.
        ''' </param>
        <ExMethod>
        Public Shared Sub AnimateAngle(formName As Primitive, controlName As Primitive, angle As Primitive, duration As Primitive)
            Try
                Dim obj = GetControl(formName, controlName)
                App.Invoke(
                    Sub()
                        GraphicsWindow.DoubleAnimateProperty(obj, AngleProperty, angle, duration)
                    End Sub)
            Catch ex As Exception
                ShowSubError(formName, controlName, "AnimateAngle", ex.Message)
            End Try
        End Sub



#End Region

#Region "Events"

        <ExMethod>
        Public Shared Sub HandleEvents(FormName As Primitive, ControlName As Primitive)
            [Event].SenderForm = FormName
            [Event].SenderControl = ControlName
        End Sub

        Shared Function GetVisualElemet(eventNmae As String) As FrameworkElement
            Try
                Dim VisualElement = CType(GetControl([Event].SenderForm, [Event].SenderControl), FrameworkElement)
                App.Invoke(
                      Sub()
                          If TypeOf VisualElement Is Window Then
                              Dim win = CType(VisualElement, Window)
                              If win.AllowsTransparency Then
                                  ' Use the camvas events ibstead of the form, because form events will not fire if it is transperent
                                  VisualElement = CType(win.Content, Wpf.Canvas)
                              End If
                          End If
                      End Sub)
                Return VisualElement
            Catch ex As Exception
                [Event].ShowErrorMessage(eventNmae, ex.Message)
            End Try
            Return Nothing
        End Function


        ''' <summary>
        ''' Fired when user presses the left mouse-button down
        ''' </summary>
        Public Shared Custom Event OnMouseLeftDown As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Dim VisualElement = GetVisualElemet(NameOf(OnMouseLeftDown))
                AddHandler VisualElement.PreviewMouseLeftButtonDown, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                Dim VisualElement = GetVisualElemet(NameOf(OnClick))
                AddHandler VisualElement.PreviewMouseLeftButtonUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                Dim VisualElement = GetVisualElemet(NameOf(OnMouseLeftUp))
                AddHandler VisualElement.PreviewMouseLeftButtonUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                Dim VisualElement = GetVisualElemet(NameOf(OnDoubleClick))
                AddHandler VisualElement.PreviewMouseLeftButtonUp,
                      Sub(Sender As Object, e As Input.MouseButtonEventArgs)
                          If e.ClickCount > 1 Then [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                      End Sub
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
                Dim VisualElement = GetVisualElemet(NameOf(OnMouseRightDown))
                AddHandler VisualElement.PreviewMouseRightButtonDown, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                Dim VisualElement = GetVisualElemet(NameOf(OnMouseRightUp))
                AddHandler VisualElement.PreviewMouseRightButtonUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                Dim VisualElement = GetVisualElemet(NameOf(OnMouseMove))
                AddHandler VisualElement.PreviewMouseMove, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                Dim VisualElement = GetVisualElemet(NameOf(OnMouseWheel))
                AddHandler VisualElement.PreviewMouseWheel, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                Dim VisualElement = GetVisualElemet(NameOf(OnMouseEnter))
                AddHandler VisualElement.MouseEnter, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                Dim VisualElement = GetVisualElemet(NameOf(OnMouseLeave))
                AddHandler VisualElement.MouseLeave, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                Dim VisualElement = GetVisualElemet(NameOf(OnKeyDown))
                AddHandler VisualElement.PreviewKeyDown, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the user presses a keyboard-ky down on a control or any of its cjild controls.
        ''' </summary>
        Public Shared Custom Event OnPreviewKeyDown As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Dim VisualElement = GetVisualElemet(NameOf(OnPreviewKeyDown))
                AddHandler VisualElement.PreviewKeyDown, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler, True)
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
                Dim VisualElement = GetVisualElemet(NameOf(OnKeyUp))
                AddHandler VisualElement.PreviewKeyUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the user releases a keyboard-ky on the control or any of its child controls.
        ''' </summary>
        Public Shared Custom Event OnPreviewKeyUp As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Dim VisualElement = GetVisualElemet(NameOf(OnPreviewKeyUp))
                AddHandler VisualElement.PreviewKeyUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler, True)
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
                Throw New Exception($"There is no form named `{formName}`.")
            End If

            Dim _controls = Forms._forms(formName)
            If controlName = "" Then Return CType(_controls(formName), Wpf.Control)

            Dim name = controlName.ToLower()
            If Not _controls.ContainsKey(name) Then
                Throw New Exception($"There is no control named `{controlName}` on form {formName}.")
            End If
            Return CType(_controls(name), Wpf.Control)
        End Function

        Shared Function GetParent(child As DependencyObject, parentType As Type) As UIElement
            If child Is Nothing Then Return Nothing
            Dim p = child
            Do
                p = VisualTreeHelper.GetParent(p)
                If p Is Nothing Then Return Nothing
                If p.GetType Is parentType Then Return CType(p, UIElement)
            Loop
        End Function
    End Class


End Namespace
