Option Explicit On


Imports Wpf = System.Windows.Controls
Imports Microsoft.SmallBasic.Library
Imports System.Windows
Imports System.Windows.Media
Imports App = Microsoft.SmallBasic.Library.Internal.SmallBasicApplication

Namespace WinForms
    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class Control

        Shared Sub ShowSubError(key As String, memberName As String, ex As Exception)
            If Not key.Contains(".") Then
                key &= "." & key
            End If
            Dim names = key.Split(".")
            ShowSubError(names(0), names(1), ex)
        End Sub


        Shared Sub ShowSubError(formName As String, controlName As String, memberName As String, ex As Exception)
            Dim msg = ex.Message
            If controlName.ToLower() = formName.ToLower Then
                ReportError($"Calling {formName}.{memberName} caused an error: {vbCrLf}{msg}", ex)
            Else
                ReportError($"Calling {formName}.{controlName}.{memberName} caused an error: {vbCrLf}{msg}", ex)
            End If
        End Sub

        Shared Sub ShowErrorMesssage(key As String, memberName As String, ex As Exception)
            If Not key.Contains(".") Then
                key &= "." & key
            End If
            Dim names = key.Split(".")
            ShowErrorMesssage(names(0), names(1), memberName, ex)
        End Sub

        Shared Sub ShowErrorMesssage(formName As String, controlName As String, memberName As String, ex As Exception)
            If controlName.ToLower() = formName.ToLower Then
                ReportError($"Reading {formName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
            Else
                ReportError($"Reading {formName}.{controlName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
            End If
        End Sub

        Shared Sub ShowPropertyMesssage(key As String, memberName As String, value As String, ex As Exception)
            If Not key.Contains(".") Then
                key &= "." & key
            End If
            Dim names = key.Split(".")
            ShowPropertyMesssage(names(0), names(1), memberName, value, ex)
        End Sub


        Shared Sub ShowPropertyMesssage(formName As String, controlName As String, memberName As String, value As String, ex As Exception)
            If controlName.ToLower() = formName.ToLower Then
                ReportError($"Sending `{value}` to {formName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
            Else
                ReportError($"Sending `{value}` to {formName}.{controlName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
            End If
        End Sub


        ''' <summary>
        ''' Gets the name of the control.           
        ''' </summary>
        ''' <remarks>You can't change the control name at runtime. Use the designer to change the name</remarks>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetName(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetName = GetControl(controlName).Name
                    Catch ex As Exception
                        ShowErrorMesssage(controlName, "Name", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' The x-pos of the control on its parent control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetLeft(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetLeft = Wpf.Canvas.GetLeft(GetControl(controlName))
                    Catch ex As Exception
                        ShowErrorMesssage(controlName, "Left", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetLeft(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
                        ' Remove any animation effect to allow setting the new value
                        obj.BeginAnimation(Wpf.Canvas.LeftProperty, Nothing)
                        Wpf.Canvas.SetLeft(obj, CDbl(value))
                    Catch ex As Exception
                        ShowPropertyMesssage(controlName, "Left", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The y-pos of the control on its parent control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetTop(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetTop = Wpf.Canvas.GetTop(GetControl(controlName))
                    Catch ex As Exception
                        ShowErrorMesssage(controlName, "Top", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetTop(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
                        ' Remove any animation effect to allow setting the new value
                        obj.BeginAnimation(Wpf.Canvas.TopProperty, Nothing)
                        Wpf.Canvas.SetTop(obj, value)
                    Catch ex As Exception
                        ShowPropertyMesssage(controlName, "Top", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The width of the control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetWidth(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
                        If TypeOf obj Is Window Then
                            GetWidth = CType(CType(obj, Window).Content, Wpf.Canvas).ActualWidth
                        Else
                            GetWidth = obj.ActualWidth
                        End If
                    Catch ex As Exception
                        ShowErrorMesssage(controlName, "Width", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetWidth(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
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
                        ShowPropertyMesssage(controlName, "Width", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The height of the control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetHeight(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
                        If TypeOf obj Is Window Then
                            GetHeight = CType(CType(obj, Window).Content, Wpf.Canvas).ActualHeight
                        Else
                            GetHeight = obj.ActualHeight
                        End If
                    Catch ex As Exception
                        ShowErrorMesssage(controlName, "Height", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetHeight(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
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
                        ShowPropertyMesssage(controlName, "Height", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' When its value = 1 (or True), user can interact with the control.
        ''' When its value = 0 (or False),  the control is disabled, and user can't interact with it.
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetEnabled(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetEnabled = GetControl(controlName).IsEnabled
                    Catch ex As Exception
                        ShowErrorMesssage(controlName, "Ebabled", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetEnabled(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetControl(controlName).IsEnabled = value
                    Catch ex As Exception
                        ShowPropertyMesssage(controlName, "Enabled", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' When its value = 1 (or True), the control is shown at the form.
        ''' When its value = 0 (or False),  the control is hidden.
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetVisible(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetVisible = GetControl(controlName).IsVisible
                    Catch ex As Exception
                        ShowErrorMesssage(controlName, "Visible", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetVisible(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetControl(controlName).Visibility = If(value, Visibility.Visible, Visibility.Hidden)
                    Catch ex As Exception
                        ShowPropertyMesssage(controlName, "Visible", value, ex)
                    End Try
                End Sub)
        End Sub


        Private Shared ReadOnly BackColorProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("BackColor",
                           GetType(Media.Color), GetType(Wpf.Control),
                           New PropertyMetadata(SystemColors.ControlColor, AddressOf BackColor_Changed))

        Private Shared Sub BackColor_Changed(c As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim _color = CType(e.NewValue, Media.Color)
            Dim brush As New SolidColorBrush(_color)
            Dim L = TryCast(c, Wpf.Label)

            If L IsNot Nothing Then
                If L.Content Is Nothing Then
                    L.Background = brush
                ElseIf TypeOf L.Content Is System.Windows.Shapes.Shape Then
                    CType(L.Content, System.Windows.Shapes.Shape).Fill = brush
                Else
                    L.Content.Background = brush
                End If
            Else
                If TypeOf c Is Window Then
                    CType(CType(c, Window).Content, Wpf.Canvas).Background = brush
                End If

                CType(c, Wpf.Control).Background = brush
            End If
        End Sub

        ''' <summary>
        ''' The backgeound color of the control.
        ''' Use values from the Color object, such as Color.Yellow
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        <ExProperty>
        Public Shared Function GetBackColor(controlName As Primitive) As Primitive
            App.Invoke(
                 Sub()
                     Try
                         Dim c = GetControl(controlName)
                         GetBackColor = GetBackColor(c).ToString()
                     Catch ex As Exception
                         ShowErrorMesssage(controlName, "BackColor", ex)
                     End Try
                 End Sub)
        End Function

        Public Shared Function GetBackColor(control As Wpf.Control) As Media.Color
            Dim brush As SolidColorBrush
            If TypeOf control Is Window Then
                Dim canvas = CType(CType(control, Window).Content, Wpf.Canvas)
                brush = TryCast(canvas.Background, SolidColorBrush)

            ElseIf TypeOf control Is Wpf.Label Then
                Dim L = CType(control, Wpf.Label)
                If L.Content Is Nothing Then
                    brush = TryCast(control.Background, SolidColorBrush)
                ElseIf TypeOf L.Content Is System.Windows.Shapes.Shape Then
                    brush = CType(L.Content, System.Windows.Shapes.Shape).Fill
                Else
                    brush = L.Content.Background
                End If

            Else
                brush = TryCast(control.Background, SolidColorBrush)
            End If

            If brush IsNot Nothing Then
                Return brush.Color
            Else
                Return CType(control.GetValue(BackColorProperty), Media.Color)
            End If
        End Function

        <ExProperty>
        Public Shared Sub SetBackColor(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim _color = Color.FromString(value)
                        Dim obj = GetControl(controlName)
                        ' Remove any animation effect to allow setting the new value
                        obj.BeginAnimation(BackColorProperty, Nothing)
                        obj.SetValue(BackColorProperty, _color)
                    Catch ex As Exception
                        ShowPropertyMesssage(controlName, "BackColor", value, ex)
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
        <ReturnValueType(VariableType.Color)>
        <ExProperty>
        Public Shared Function GetForeColor(controlName As Primitive) As Primitive
            App.Invoke(
                  Sub()
                      Try
                          Dim c = GetControl(controlName)
                          Dim brush = TryCast(c.Foreground, SolidColorBrush)
                          If brush IsNot Nothing Then
                              GetForeColor = brush.Color.ToString()
                          Else
                              GetForeColor = CStr(c.GetValue(ForeColorProperty))
                          End If
                      Catch ex As Exception
                          ShowErrorMesssage(controlName, "ForeColor", ex)
                      End Try
                  End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetForeColor(controlName As Primitive, value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Dim c = GetControl(controlName)
                           Dim _color = Color.FromString(value)
                           c.Foreground = New SolidColorBrush(_color)
                           c.SetValue(ForeColorProperty, value.ToString())
                       Catch ex As Exception
                           ShowPropertyMesssage(controlName, "ForeColor", value, ex)
                       End Try
                   End Sub)
        End Sub

        ''' <summary>
        ''' The mouse x-pos relative to the control. When mouse is over the control, this value lies between 0 and the control's width.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetMouseX(controlName As Primitive) As Primitive
            App.Invoke(
                   Sub()
                       Try
                           Dim c = GetControl(controlName)
                           GetMouseX = System.Math.Round(Input.Mouse.GetPosition(c).X)
                       Catch ex As Exception
                           ShowErrorMesssage(controlName, "MouseX", ex)
                       End Try
                   End Sub)
        End Function

        ''' <summary>
        ''' The mouse y-pos relative to the control. When mouse is over the control, this value lies between 0 and the control's.height.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetMouseY(controlName As Primitive) As Primitive
            App.Invoke(
                    Sub()
                        Try
                            Dim c = GetControl(controlName)
                            GetMouseY = System.Math.Round(Input.Mouse.GetPosition(c).Y)
                        Catch ex As Exception
                            ShowErrorMesssage(controlName, "MouseY", ex)
                        End Try
                    End Sub)
        End Function


        ''' <summary>
        ''' Gets or sets the rotation angle of the control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetAngle(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetAngle = GetAngle(GetControl(controlName))
                    Catch ex As Exception
                        ShowErrorMesssage(controlName, "Angle", ex)
                    End Try
                End Sub)
        End Function

        Private Shared Function GetAngle(element As DependencyObject) As Double
            If element Is Nothing Then
                Throw New ArgumentNullException("element")
            End If

            Return CDbl(element.GetValue(AngleProperty))
        End Function


        <ExProperty>
        Public Shared Sub SetAngle(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        SetAngle(GetControl(controlName), value)
                    Catch ex As Exception
                        ShowPropertyMesssage(controlName, "Angle", value, ex)
                    End Try
                End Sub)
        End Sub

        Private Shared Sub SetAngle(element As DependencyObject, value As Double)
            If element Is Nothing Then
                Throw New ArgumentNullException("element")
            End If

            CType(element, Wpf.Control).BeginAnimation(AngleProperty, Nothing)
            element.SetValue(AngleProperty, value)
        End Sub


        Private Shared _rotateTransformMap As New Dictionary(Of Wpf.Control, RotateTransform)

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

        ''' <summary>
        ''' Moves focus to the control, so it beccomes the active control that recives the keybored keys.
        ''' </summary>
        <ExMethod>
        Public Shared Sub Focus(controlName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetControl(controlName).Focus()
                    Catch ex As Exception
                        ShowSubError(controlName, "Focus", ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Rotates the control by the specified angle.
        ''' </summary>
        ''' <param name="angle">
        ''' The angle to rotate the shape by. It will be added to the shape current angle.
        ''' </param>
        <ExMethod>
        Public Shared Sub Rotate(controlName As Primitive, angle As Primitive)
            Try
                Dim obj = GetControl(controlName)
                App.BeginInvoke(Sub() SetAngle(obj, GetAngle(obj) + CDbl(angle)))
            Catch ex As Exception
                ShowSubError(controlName, "Rotate", ex)
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
        Public Shared Sub AnimateColor(controlName As Primitive, newColor As Primitive, duration As Primitive)
            Try
                Dim obj = GetControl(controlName)
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
                ShowSubError(controlName, "AnimateColor", ex)
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
        Public Shared Sub AnimateTransparency(controlName As Primitive, transparency As Primitive, duration As Primitive)
            Try
                Dim c = GetBackColor(controlName)
                c = Color.SetTransparency(c, transparency)
                AnimateColor(controlName, c, duration)
            Catch ex As Exception
                ShowSubError(controlName, "AnimateColor", ex)
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
        Public Shared Sub AnimatePos(controlName As Primitive, x As Primitive, y As Primitive, duration As Primitive)
            Try
                Dim obj = GetControl(controlName)
                App.Invoke(
                    Sub()
                        GraphicsWindow.DoubleAnimateProperty(obj, Wpf.Canvas.LeftProperty, x, CDbl(duration))
                        GraphicsWindow.DoubleAnimateProperty(obj, Wpf.Canvas.TopProperty, y, CDbl(duration))
                    End Sub)
            Catch ex As Exception
                ShowSubError(controlName, "AnimatePos", ex)
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
        Public Shared Sub AnimateSize(controlName As Primitive, width As Primitive, height As Primitive, duration As Primitive)
            Try
                Dim obj = GetControl(controlName)
                App.Invoke(
                    Sub()
                        GraphicsWindow.DoubleAnimateProperty(obj, FrameworkElement.WidthProperty, width, duration)
                        GraphicsWindow.DoubleAnimateProperty(obj, FrameworkElement.HeightProperty, height, duration)
                    End Sub)
            Catch ex As Exception
                ShowSubError(controlName, "AnimateSize", ex)
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
        Public Shared Sub AnimateAngle(controlName As Primitive, angle As Primitive, duration As Primitive)
            Try
                Dim obj = GetControl(controlName)
                App.Invoke(
                    Sub()
                        GraphicsWindow.DoubleAnimateProperty(obj, AngleProperty, CDbl(angle), CDbl(duration))
                    End Sub)
            Catch ex As Exception
                ShowSubError(controlName, "AnimateAngle", ex)
            End Try
        End Sub

#End Region

#Region "Events"

        <ExMethod>
        Public Shared Sub HandleEvents(ControlName As Primitive)
            [Event].SenderControl = ControlName
        End Sub

        Shared Function GetVisualElemet(eventNmae As String, Optional getForm As Boolean = False) As FrameworkElement
            Try
                Dim VisualElement = CType(GetControl([Event].SenderControl), FrameworkElement)
                If TypeOf VisualElement Is Window AndAlso Not getForm Then
                    App.Invoke(
                           Sub()
                               Dim win = CType(VisualElement, Window)
                               Dim canvas = CType(win.Content, Wpf.Canvas)
                               If canvas.Background IsNot Nothing OrElse win.AllowsTransparency Then
                                   ' Use the camvas events instead of the form, because form events will not fire if the form is transperent or if the canvas has a non-transparent back color
                                   VisualElement = canvas
                               End If
                           End Sub)
                End If
                Return VisualElement

            Catch ex As Exception
                [Event].ShowErrorMessage(eventNmae, ex)
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
                AddHandler VisualElement.PreviewMouseLeftButtonUp,
                    Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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

                If TypeOf VisualElement Is Wpf.Canvas Then
                    AddHandler VisualElement.MouseLeftButtonDown,
                        Sub(Sender As Object, e As Input.MouseButtonEventArgs)
                            If e.ClickCount > 1 Then [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                Else
                    AddHandler VisualElement.PreviewMouseLeftButtonDown,
                         Sub(Sender As Object, e As Input.MouseButtonEventArgs)
                             If e.ClickCount > 1 Then [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                         End Sub
                End If

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
                Dim VisualElement = GetVisualElemet(NameOf(OnKeyDown), True)
                If TypeOf VisualElement Is Window OrElse TypeOf VisualElement Is Wpf.Canvas Then
                    ' The form should not preview keys, to let controls handle them
                    AddHandler VisualElement.KeyDown,
                        Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Else
                    ' Preview keydown can habdle space and other keys im TextBox
                    AddHandler VisualElement.PreviewKeyDown,
                        Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                End If
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
                Dim VisualElement = GetVisualElemet(NameOf(OnPreviewKeyDown), True)
                AddHandler VisualElement.PreviewKeyDown,
                    Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler, True)
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
                Dim VisualElement = GetVisualElemet(NameOf(OnKeyUp), True)
                If TypeOf VisualElement Is Window OrElse TypeOf VisualElement Is Wpf.Canvas Then
                    AddHandler VisualElement.KeyUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Else
                    AddHandler VisualElement.PreviewKeyUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                End If
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
                Dim VisualElement = GetVisualElemet(NameOf(OnPreviewKeyUp), True)
                AddHandler VisualElement.PreviewKeyUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler, True)
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event
#End Region

        Friend Shared Function GetControl(key As String) As Wpf.Control
            Return GetFrameworkElement(key)
        End Function

        Friend Shared Function GetFrameworkElement(key As String) As FrameworkElement
            key = key.ToLower()
            If Not key.Contains(".") Then
                key &= "." & key
            End If

            If Not Forms._controls.ContainsKey(key) Then
                Dim names = key.Split("."c)
                Throw New Exception($"There is no control named `{names(1)}` on form {names(0)}.")
            End If

            Return CType(Forms._controls(key), Wpf.Control)
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