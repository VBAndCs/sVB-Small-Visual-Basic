
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Markup
Imports System.Windows.Media
Imports Microsoft.SmallVisualBasic.Library
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports Wpf = System.Windows.Controls

Namespace WinForms
    ''' <summary>
    ''' This is the parent control of other controls like Form, Label, and Button controls. All controls inherit the properties, methods and events defined in this type.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class Control

        Shared Sub ReportSubError(key As String, memberName As String, ex As Exception)
            Dim names = If(key.Contains("."),
                key.Split("."c),
                {key, key}
            )

            ReportSubError(names(0), names(1), memberName, ex)
        End Sub

        Shared Sub ReportSubError(formName As String, controlName As String, memberName As String, ex As Exception)
            Dim msg = ex.Message
            If controlName.ToLower() = formName.ToLower Then
                Helper.ReportError($"Calling {formName}.{memberName} caused an error: {vbCrLf}{msg}", ex)
            Else
                Helper.ReportError($"Calling {formName}.{controlName}.{memberName} caused an error: {vbCrLf}{msg}", ex)
            End If
        End Sub

        Shared Sub ReportError(key As String, memberName As String, ex As Exception)
            Dim names = If(key.Contains("."),
                key.Split("."c),
                {key, key}
            )
            ReportError(names(0), names(1), memberName, ex)
        End Sub

        Private Shared Sub ReportError(formName As String, controlName As String, memberName As String, ex As Exception)
            If controlName.ToLower() = formName.ToLower Then
                Helper.ReportError($"Reading {formName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
            Else
                Helper.ReportError($"Reading {formName}.{controlName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
            End If
        End Sub

        Shared Sub ReportPropertyError(key As String, memberName As String, value As String, ex As Exception)
            Dim names = If(key.Contains("."),
                key.Split("."c),
                {key, key}
            )
            ReportyPropertyError(names(0), names(1), memberName, value, ex)
        End Sub


        Shared Sub ReportyPropertyError(formName As String, controlName As String, memberName As String, value As String, ex As Exception)
            If controlName.ToLower() = formName.ToLower Then
                Helper.ReportError($"Sending `{value}` to {formName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
            Else
                Helper.ReportError($"Sending `{value}` to {formName}.{controlName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
            End If
        End Sub


        ''' <summary>
        ''' Gets the name of the control.
        ''' Note that you can't change the control name at runtime. Use the designer to change the name.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetName(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetName = New Primitive(GetControl(controlName).Name)
                    Catch ex As Exception
                        ReportError(controlName, "Name", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Gets the form that contains the given control. 
        ''' If the control is a form, it will be the return value.
        ''' If the control has no parent form, the return value will be an empty string "".
        ''' </summary>
        <ReturnValueType(VariableType.Form)>
        <ExProperty>
        Public Shared Function GetParentForm(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim name = controlName.AsString().ToLower()
                        Dim i = name.IndexOf(".")
                        If i = -1 Then
                            If Forms._forms.ContainsKey(name) Then
                                GetParentForm = New Primitive(name)
                            Else
                                GetParentForm = New Primitive("")
                            End If
                        Else
                            GetParentForm = New Primitive(name.Substring(0, i))
                        End If

                    Catch ex As Exception
                        ReportError(controlName, "Name", ex)
                    End Try
                End Sub)
        End Function


        ''' <summary>
        ''' Gets the name of the control type, like TextBox, Label, and Button.           
        ''' </summary>
        <ReturnValueType(VariableType.ControlType)>
        <ExProperty>
        Public Shared Function GetTypeName(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        If WinTimer.Timers.ContainsKey(controlName) Then
                            GetTypeName = ControlTypes.WinTimer
                        Else
                            GetTypeName = New Primitive(GetControl(controlName).GetType().Name)
                        End If
                    Catch ex As Exception
                        ReportError(controlName, "TypeName", ex)
                    End Try
                End Sub)
        End Function


        ''' <summary>
        ''' An extra property to store any value you want that is related to the control. you can store multiple values as by putting them in an array or a dynamic object
        ''' </summary>
        <ExProperty>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetTag(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim x = New Primitive(CStr(GetControl(controlName).Tag))
                        x.ConstructArrayMap()
                        GetTag = x
                    Catch ex As Exception
                        ReportError(controlName, "Tag", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetTag(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetControl(controlName).Tag = value.AsString()
                    Catch ex As Exception
                        ReportPropertyError(controlName, "Tag", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the x-pos of the left edge of control on its parent control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetLeft(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetLeft = System.Math.Round(Wpf.Canvas.GetLeft(GetControl(controlName)), 2, MidpointRounding.AwayFromZero)
                    Catch ex As Exception
                        ReportError(controlName, "Left", ex)
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
                        ReportPropertyError(controlName, "Left", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the x-pos of the right edge of control on its parent control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetRight(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)
                        GetRight = System.Math.Round(Wpf.Canvas.GetLeft(c) + c.ActualWidth, 2, MidpointRounding.AwayFromZero)
                    Catch ex As Exception
                        ReportError(controlName, "Right", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetRight(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)
                        ' Remove any animation effect to allow setting the new value
                        c.BeginAnimation(Wpf.Canvas.LeftProperty, Nothing)
                        Wpf.Canvas.SetLeft(c, CDbl(value) - c.ActualWidth)
                    Catch ex As Exception
                        ReportPropertyError(controlName, "Right", value, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Gets or sets the y-pos of the top edge of the control on its parent control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetTop(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetTop = System.Math.Round(Wpf.Canvas.GetTop(GetControl(controlName)), 2, MidpointRounding.AwayFromZero)
                    Catch ex As Exception
                        ReportError(controlName, "Top", ex)
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
                        Wpf.Canvas.SetTop(obj, CDbl(value))
                    Catch ex As Exception
                        ReportPropertyError(controlName, "Top", value, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Gets or sets the y-pos of the bottom edge of the control on its parent control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetBottom(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)
                        GetBottom = System.Math.Round(Wpf.Canvas.GetTop(c) + c.ActualWidth, 2, MidpointRounding.AwayFromZero)
                    Catch ex As Exception
                        ReportError(controlName, "Bottom", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetBottom(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)
                        ' Remove any animation effect to allow setting the new value
                        c.BeginAnimation(Wpf.Canvas.TopProperty, Nothing)
                        Wpf.Canvas.SetTop(c, CDbl(value) - c.ActualHeight)
                    Catch ex As Exception
                        ReportPropertyError(controlName, "Bottom", value, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Gets or sets the actual x-pos of the rotated control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetRotatedLeft(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim ctrl = GetControl(controlName)
                        Dim left = Canvas.GetLeft(ctrl)
                        Dim rotation = TryCast(ctrl.RenderTransform, RotateTransform)
                        If rotation Is Nothing Then
                            GetRotatedLeft = System.Math.Round(left, 2, MidpointRounding.AwayFromZero)
                            Return
                        End If

                        Dim c = ctrl.RenderTransformOrigin
                        Dim cx = ctrl.ActualWidth * c.X
                        Dim cy = ctrl.ActualHeight * c.Y
                        Dim angle = rotation.Angle * (Math.Pi / 180)
                        Dim offsetX = cx * (1 - Math.Cos(angle)) + cy * Math.Sin(angle)
                        GetRotatedLeft = System.Math.Round(offsetX + left, 2, MidpointRounding.AwayFromZero)

                    Catch ex As Exception
                        ReportError(controlName, "RotatedLeft", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetRotatedLeft(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim ctrl = GetControl(controlName)
                        Dim offsetX = 0
                        ' Remove any animation effect to allow setting the new value
                        ctrl.BeginAnimation(Wpf.Canvas.LeftProperty, Nothing)

                        Dim rotation = TryCast(ctrl.RenderTransform, RotateTransform)
                        If rotation IsNot Nothing Then
                            Dim c = ctrl.RenderTransformOrigin
                            Dim cx = ctrl.ActualWidth * c.X
                            Dim cy = ctrl.ActualHeight * c.Y
                            Dim angle = rotation.Angle * (Math.Pi / 180)
                            offsetX = cx * (1 - Math.Cos(angle)) + cy * Math.Sin(angle)
                        End If
                        Canvas.SetLeft(ctrl, CDbl(value) - offsetX)

                    Catch ex As Exception
                        ReportPropertyError(controlName, "RotatedLeft", value, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Gets or sets the actual y-pos of the rotated control.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetRotatedTop(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim ctrl = GetControl(controlName)
                        Dim top = Wpf.Canvas.GetTop(ctrl)
                        Dim rotation = TryCast(ctrl.RenderTransform, RotateTransform)
                        If rotation Is Nothing Then
                            GetRotatedTop = System.Math.Round(top, 2, MidpointRounding.AwayFromZero)
                            Return
                        End If

                        Dim c = ctrl.RenderTransformOrigin
                        Dim cx = ctrl.ActualWidth * c.X
                        Dim cy = ctrl.ActualHeight * c.Y
                        Dim angle = rotation.Angle * (Math.Pi / 180)
                        Dim offsetY = cy * (1 - Math.Cos(angle)) - cx * Math.Sin(angle)
                        GetRotatedTop = System.Math.Round(offsetY + top, 2, MidpointRounding.AwayFromZero)

                    Catch ex As Exception
                        ReportError(controlName, "RotatedTop", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetRotatedTop(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim ctrl = GetControl(controlName)
                        Dim offsetY = 0
                        ' Remove any animation effect to allow setting the new value
                        ctrl.BeginAnimation(Wpf.Canvas.TopProperty, Nothing)

                        Dim rotation = TryCast(ctrl.RenderTransform, RotateTransform)
                        If rotation IsNot Nothing Then
                            Dim c = ctrl.RenderTransformOrigin
                            Dim cx = ctrl.ActualWidth * c.X
                            Dim cy = ctrl.ActualHeight * c.Y
                            Dim angle = rotation.Angle * (Math.Pi / 180)
                            offsetY = cy * (1 - Math.Cos(angle)) - cx * Math.Sin(angle)
                        End If
                        Canvas.SetTop(ctrl, CDbl(value) - offsetY)

                    Catch ex As Exception
                        ReportPropertyError(controlName, "RotatedTop", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the width of the control.
        ''' If you want an auto-width to fit the control's content width, set this property to 0 or amy negative number. 
        ''' You can also limit the auto width by setting the MaxWidth property.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetWidth(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
                        If TypeOf obj Is Window Then
                            Dim canvas = Form.GetCanvas(CType(obj, Window))
                            GetWidth = System.Math.Round(canvas.ActualWidth, 2, MidpointRounding.AwayFromZero)
                        Else
                            GetWidth = System.Math.Round(obj.ActualWidth, 2, MidpointRounding.AwayFromZero)
                        End If

                    Catch ex As Exception
                        ReportError(controlName, "Width", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetWidth(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim w = CDbl(value)
                        If w <= 0 Then w = Double.NaN
                        Dim obj = GetControl(controlName)

                        If TypeOf obj Is Window Then
                            ' Remove any animation effect to allow setting the new value
                            obj.BeginAnimation(FrameworkElement.WidthProperty, Nothing)
                            Form.GetCanvas(obj).Width = w
                        Else
                            ' Remove any animation effect to allow setting the new value
                            obj.BeginAnimation(FrameworkElement.WidthProperty, Nothing)
                            obj.Width = w
                        End If

                    Catch ex As Exception
                        ReportPropertyError(controlName, "Width", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the height of the control. 
        ''' If you want an auto-heigh to fit the control's content heigh, set this property to 0 or any negative number. 
        ''' You can also limit the auto height by setting the MaxHeigh property.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetHeight(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
                        If TypeOf obj Is Window Then
                            GetHeight = System.Math.Round(Form.GetCanvas(obj).ActualHeight, 2, MidpointRounding.AwayFromZero)
                        Else
                            GetHeight = System.Math.Round(obj.ActualHeight, 2, MidpointRounding.AwayFromZero)
                        End If

                    Catch ex As Exception
                        ReportError(controlName, "Height", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetHeight(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim h = CDbl(value)
                        If h <= 0 Then h = Double.NaN
                        Dim obj = GetControl(controlName)

                        If TypeOf obj Is Window Then
                            ' Remove any animation effect to allow setting the new value
                            obj.BeginAnimation(FrameworkElement.HeightProperty, Nothing)
                            Form.GetCanvas(obj).Height = h
                        Else
                            ' Remove any animation effect to allow setting the new value
                            obj.BeginAnimation(FrameworkElement.HeightProperty, Nothing)
                            obj.Height = h
                        End If

                    Catch ex As Exception
                        ReportPropertyError(controlName, "Height", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The max width of the control. It is useful especially when you set the control's width to auto (Width = -1), to force the auto witdth not to exceed a max length.
        ''' Setting this property to 0 or a negative value will reset it (no max width).
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetMaxWidth(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
                        Dim value As Double

                        If TypeOf obj Is Window Then
                            value = Form.GetCanvas(obj).MaxWidth
                        Else
                            value = obj.MaxWidth
                        End If

                        If Double.IsPositiveInfinity(value) Then
                            GetMaxWidth = 0
                        Else
                            GetMaxWidth = System.Math.Round(value, 2, MidpointRounding.AwayFromZero)
                        End If

                    Catch ex As Exception
                        ReportError(controlName, "MaxWidth", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetMaxWidth(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim w = CDbl(value)
                        If w <= 0 Then w = Double.PositiveInfinity
                        Dim obj = GetControl(controlName)

                        If TypeOf obj Is Window Then
                            Form.GetCanvas(obj).MaxWidth = w
                        Else
                            obj.MaxWidth = w
                        End If

                    Catch ex As Exception
                        ReportPropertyError(controlName, "MaxWidth", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' The max height of the control. It is useful especially when you set the control's height to auto (Height = -1), to force the auto height not to exceed a max length.
        ''' Setting this property to 0 or a negative value will reset it (no max height).
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetMaxHeight(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
                        Dim value As Double

                        If TypeOf obj Is Window Then
                            value = Form.GetCanvas(obj).MaxHeight
                        Else
                            value = obj.MaxHeight
                        End If

                        If Double.IsPositiveInfinity(value) Then
                            GetMaxHeight = 0
                        Else
                            GetMaxHeight = System.Math.Round(value, 2, MidpointRounding.AwayFromZero)
                        End If

                    Catch ex As Exception
                        ReportError(controlName, "MaxHeight", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetMaxHeight(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim h = CDbl(value)
                        If h <= 0 Then h = Double.PositiveInfinity
                        Dim obj = GetControl(controlName)

                        If TypeOf obj Is Window Then
                            Form.GetCanvas(obj).MaxHeight = h
                        Else
                            obj.MaxHeight = h
                        End If

                    Catch ex As Exception
                        ReportPropertyError(controlName, "MaxHeight", value, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Gets or sets the padding of the control, which is the internal space between the control edges and its content.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetPadding(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetPadding = System.Math.Round(GetControl(controlName).Padding.Left, 1, MidpointRounding.AwayFromZero)
                    Catch ex As Exception
                        ReportError(controlName, "Padding", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetPadding(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetControl(controlName).Padding = New Thickness(CDbl(value))
                    Catch ex As Exception
                        ReportPropertyError(controlName, "Padding", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Changes the witdth of the control to fit its content width.
        ''' This is a one time change, and will not make the control width auto-szied.
        ''' If you want to make the width auto-sized, set the Width property to -1.
        ''' </summary>
        <ExMethod>
        Public Shared Sub FitContentWidth(controlName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)
                        c.Width = Double.NaN
                        c.Dispatcher.Invoke(Threading.DispatcherPriority.Render, Sub() Exit Sub)
                        c.Width = c.ActualWidth

                    Catch ex As Exception
                        ReportSubError(controlName, "FitContentWidth", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Changes the height of the control to fit its content height. This may be useful when you set WordWrap = True in labels and buttons.
        ''' This is a one time change, and will not make the control height auto-szied.
        ''' If you want to make the height auto-sized, set the Height property to -1.
        ''' </summary>
        <ExMethod>
        Public Shared Sub FitContentHeight(controlName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)
                        c.Height = Double.NaN
                        c.Dispatcher.Invoke(Threading.DispatcherPriority.Render, Sub() Exit Sub)
                        c.Height = c.ActualHeight

                    Catch ex As Exception
                        ReportSubError(controlName, "FitContentHeight", ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Changes the witdth and height of the control to fit its content size.
        ''' This is a one time change, and will not make the control width and height auto-szied.
        ''' If you want to make them auto-sized, set both Width and Height properties to -1.
        ''' </summary>
        <ExMethod>
        Public Shared Sub FitContentSize(controlName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)
                        c.Width = Double.NaN
                        c.Height = Double.NaN
                        c.Dispatcher.Invoke(Threading.DispatcherPriority.Render, Sub() Exit Sub)
                        c.Width = c.ActualWidth
                        c.Height = c.ActualHeight

                    Catch ex As Exception
                        ReportSubError(controlName, "FitContentSize", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' When it is True, user can interact with the control.
        ''' When it is False,  the control is disabled, and user can't interact with it.
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetEnabled(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetEnabled = GetControl(controlName).IsEnabled
                    Catch ex As Exception
                        ReportError(controlName, "Enabled", ex)
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
                        ReportPropertyError(controlName, "Enabled", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the visibilty state of the control:
        ''' When it is True, the control is displayed on the form.
        ''' When it is False,  the control is hidden.
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetVisible(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetVisible = GetControl(controlName).IsVisible
                    Catch ex As Exception
                        ReportError(controlName, "Visible", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetVisible(controlName As Primitive, value As Primitive)
            If controlName.IsEmpty Then Stop
            App.Invoke(
                Sub()
                    Try
                        Dim ctrl = GetControl(controlName)
                        If TypeOf ctrl Is Wpf.Label Then
                            Dim _label = CType(ctrl, Wpf.Label)
                            Dim tb = TryCast(_label.Content, Wpf.TextBlock)
                            If tb Is Nothing Then
                                _label.Visibility = If(value, Visibility.Visible, Visibility.Hidden)
                            ElseIf CBool(value) Then
                                _label.Visibility = Visibility.Visible
                                tb.Visibility = Visibility.Visible
                            Else
                                _label.Visibility = Visibility.Hidden
                                tb.Visibility = Visibility.Collapsed
                            End If
                        Else
                            ctrl.Visibility = If(value, Visibility.Visible, Visibility.Hidden)
                        End If
                    Catch ex As Exception
                        ReportPropertyError(controlName, "Visible", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets whether the control content is displayed from right to left.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetRightToLeft(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetRightToLeft = (GetControl(controlName).FlowDirection = FlowDirection.RightToLeft)
                    Catch ex As Exception
                        ReportError(controlName, "RightToLeft", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetRightToLeft(controlName As Primitive, value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Dim c = GetControl(controlName)
                           Dim rtl = If(CBool(value), FlowDirection.RightToLeft, FlowDirection.LeftToRight)
                           c.FlowDirection = rtl
                           If TypeOf c Is Window Then Form.GetCanvas(c).FlowDirection = rtl
                       Catch ex As Exception
                           ReportPropertyError(controlName, "RightToLeft", value, ex)
                       End Try
                   End Sub)
        End Sub


        ''' <summary>
        ''' Gets or sets the culture that the control will use to show its content and format its text.
        ''' Valid values are in the form en-US (for English Uinted States), ar-Eg (for Arabic Egypt)... etc 
        ''' </summary>
        <ExProperty>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetLanguage(ControlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim ctrl = GetControl(ControlName)
                        GetLanguage = ctrl.Language.ToString()
                    Catch ex As Exception
                        ReportError(ControlName, "Language", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetLanguage(ControlName As Primitive, cultureName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim ctrl = GetControl(ControlName)
                        ctrl.Language = XmlLanguage.GetLanguage(cultureName.AsString())
                    Catch ex As Exception
                        ReportPropertyError(ControlName, "Language", cultureName, ex)
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
                        Dim pos = System.Windows.Forms.Cursor.Position
                        GetMouseX = System.Math.Round(
                              c.PointFromScreen(New Point(pos.X, pos.Y)).X,
                              MidpointRounding.AwayFromZero)

                    Catch ex As Exception
                        ReportError(controlName, "MouseX", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' The mouse y-pos relative to the control. When mouse is over the control, this value lies between 0 and the control's height.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetMouseY(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)
                        Dim pos = System.Windows.Forms.Cursor.Position
                        GetMouseY = System.Math.Round(
                              c.PointFromScreen(New Point(pos.X, pos.Y)).Y,
                              MidpointRounding.AwayFromZero)

                    Catch ex As Exception
                        ReportError(controlName, "MouseY", ex)
                    End Try
                End Sub)
        End Function

        Private Shared focusTimes As Integer

        ''' <summary>
        ''' Moves focus to the control, so it beccomes the active control that recives the keybored keys.
        ''' </summary>
        <ExMethod>
        Public Shared Sub Focus(controlName As Primitive)
            App.Invoke(
                 Sub()
                     Try
                         Dim c = GetControl(controlName)
                         If Not c.IsFocused Then c.Focus()

                     Catch ex As Exception
                         ReportSubError(controlName, "Focus", ex)
                     End Try
                 End Sub)
        End Sub

        ''' <summary>
        ''' Releases the mouse if the currrent control captured it by calling the CaptureMouse method.
        ''' </summary>
        <ExMethod>
        Public Shared Sub ReleaseMouse(controlName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetControl(controlName).ReleaseMouseCapture()
                    Catch ex As Exception
                        ReportSubError(controlName, "ReleaseMouse", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Makes the current control owns the mouse and its events, until you call the ReleaseMouse function.
        ''' </summary>
        <ExMethod>
        Public Shared Sub CaptureMouse(controlName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Input.Mouse.Capture(GetControl(controlName), Input.CaptureMode.SubTree)
                    Catch ex As Exception
                        ReportSubError(controlName, "CaptureMouse", ex)
                    End Try
                End Sub)
        End Sub

        Public Shared ReadOnly ErrorProperty As _
                   DependencyProperty = DependencyProperty.RegisterAttached("Error",
                   GetType(String), GetType(UIElement),
                   New PropertyMetadata(""))

        Public Shared ReadOnly TipProperty As _
                   DependencyProperty = DependencyProperty.RegisterAttached("Tip",
                   GetType(String), GetType(UIElement),
                   New PropertyMetadata(Nothing))

        Public Shared ReadOnly BorderThicknessProperty As _
                   DependencyProperty = DependencyProperty.RegisterAttached("BorderThickness",
                   GetType(Thickness), GetType(UIElement),
                   New PropertyMetadata(Nothing))

        Public Shared ReadOnly BorderBrushProperty As _
                   DependencyProperty = DependencyProperty.RegisterAttached("BorderBrush",
                   GetType(Brush), GetType(UIElement),
                   New PropertyMetadata(Nothing))

        ''' <summary>
        ''' Gets or sets the error message for this control. 
        ''' When you set the error message, the control will display a red border, And the error message will shown as a tooltip when the mouse hovers over the control.
        ''' To reset the error, jsut set this property to an empty string.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetError(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)
                        GetError = New Primitive(CStr(c.GetValue(ErrorProperty)))
                    Catch ex As Exception
                        ReportError(controlName, "Error", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetError(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim errMsg = value.AsString()
                        Dim c = GetControl(controlName)
                        c.SetValue(ErrorProperty, errMsg)

                        If errMsg = "" Then
                            c.BorderThickness = c.GetValue(BorderThicknessProperty)
                            c.BorderBrush = c.GetValue(BorderBrushProperty)
                            c.ToolTip = c.GetValue(TipProperty)
                        Else
                            c.BorderThickness = New Thickness(2)
                            c.BorderBrush = Brushes.Red
                            c.ToolTip = errMsg
                        End If

                    Catch ex As Exception
                        ReportPropertyError(controlName, "Error", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the message to display as a tooltip when you hover over the control. 
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetToolTip(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)
                        GetToolTip = New Primitive(If(c.GetValue(TipProperty), "").ToString())
                    Catch ex As Exception
                        ReportError(controlName, "ToolTip", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetToolTip(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)
                        Dim x = If(value.IsEmpty, Nothing, value.AsString())
                        c.SetValue(TipProperty, x)
                        c.ToolTip = x

                    Catch ex As Exception
                        ReportPropertyError(controlName, "ToolTip", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Sets the resource dictionary of the control from the given file. The resource dictionary contains styles for the control and its child controls. This allows you to design advanced styles with other tools like VS.NET, save them to a file as a resource dictionary, then use this method to load it.
        ''' </summary>
        ''' <param name="fileName">The xaml file that contains the resource dictionary.</param>
        <ExMethod>
        Public Shared Sub SetResourceDictionary(controlName As Primitive, fileName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim ctrl = GetControl(controlName)
                        If fileName.IsEmpty Then
                            ctrl.Resources.Clear()
                            Return
                        End If

                        Dim file = If(
                             IO.Path.IsPathRooted(fileName),
                             fileName.AsString(),
                             IO.Path.Combine(Program.Directory, fileName)
                        )

                        Dim xaml = IO.File.ReadAllText(file, System.Text.Encoding.UTF8)
                        xaml = Forms.ExpandRelativeImageFiles(xaml, file)
                        Dim stream = New IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(xaml))
                        Dim resDic = CType(Markup.XamlReader.Load(stream), ResourceDictionary)
                        ctrl.Resources = resDic

                    Catch ex As Exception
                        ReportSubError(controlName, "SetResourceDictionary", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Sets the style of the control by loading it from the given file.
        ''' This allows you To design advanced styles With other tools Like VS.NET, save them To a file As a resource dictionary, Then use this method To load it.
        ''' </summary>
        ''' <param name="fileName">The xaml file that contains the resource dictionary.</param>
        ''' <param name="styleKey">The key of style. The resource dictionary can contain many styles targetting the same control. This method will search for the style that have this key.</param>
        <ExMethod>
        Public Shared Sub SetStyle(controlName As Primitive, fileName As Primitive, styleKey As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim ctrl = GetControl(controlName)
                        If fileName.IsEmpty Then
                            ctrl.Style = Nothing
                            Return
                        End If

                        Dim file = If(
                             IO.Path.IsPathRooted(fileName),
                             fileName.AsString(),
                             IO.Path.Combine(Program.Directory, fileName)
                         )
                        Dim xaml = IO.File.ReadAllText(file, System.Text.Encoding.UTF8)
                        xaml = Forms.ExpandRelativeImageFiles(xaml, file)
                        Dim stream = New IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(xaml))
                        Dim resDic = CType(Markup.XamlReader.Load(stream), ResourceDictionary)
                        ctrl.Style = CType(resDic(styleKey.AsString()), Style)

                    Catch ex As Exception
                        ReportSubError(controlName, "SetStyle", styleKey, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Raises the OnLostFocus event to apply any validation logic supplied by you in the OnLostFocus handler, then checks the Error property to see if the control has errors or not.
        ''' </summary>
        ''' <returns>True if the current control has no errors, or False otherwise.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function Validate(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)

                        ' The controle may not be validated, so, force it to validate
                        If CStr(c.GetValue(ErrorProperty)) = "" Then
                            c.RaiseEvent(New RoutedEventArgs(UIElement.LostFocusEvent))
                        End If

                        ' Now check if the control has errors
                        If CStr(c.GetValue(ErrorProperty)) = "" Then
                            Validate = New Primitive(True)
                        Else
                            Validate = New Primitive(False)
                        End If
                    Catch ex As Exception
                        ReportSubError(controlName, "Validate", ex)
                    End Try
                End Sub)
        End Function

#Region "Color and font"

        Private Shared ReadOnly BackColorProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("BackColor",
                           GetType(Media.Color?), GetType(Wpf.Control),
                           New PropertyMetadata(SystemColors.ControlColor, AddressOf BackColor_Changed))

        Private Shared Sub BackColor_Changed(c As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim brush As SolidColorBrush = Nothing
            If e.NewValue IsNot Nothing Then
                Dim _color = CType(e.NewValue, Media.Color)
                brush = New SolidColorBrush(_color)
            End If

            Dim L = TryCast(c, Wpf.Label)
            If L IsNot Nothing Then
                If TypeOf L.Content Is System.Windows.Shapes.Shape Then
                    L.Background = Nothing
                    CType(L.Content, System.Windows.Shapes.Shape).Fill = brush
                Else
                    L.Background = brush
                End If

            Else
                CType(c, Wpf.Control).Background = brush
                If TypeOf c Is Window Then
                    Form.GetCanvas(c).Background = brush
                End If
            End If
        End Sub

        ''' <summary>
        ''' The backgeound color of the control.
        ''' Use values from the Colors object, such as Colors.Yellow
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        <ExProperty>
        Public Shared Function GetBackColor(controlName As Primitive) As Primitive
            App.Invoke(
                 Sub()
                     Try
                         Dim c = GetControl(controlName)
                         GetBackColor = New Primitive(Color.GetHexaName(GetBackColor(c)))
                     Catch ex As Exception
                         ReportError(controlName, "BackColor", ex)
                     End Try
                 End Sub)
        End Function

        Public Shared Function GetBackColor(control As Wpf.Control) As Media.Color?
            Dim brush As SolidColorBrush
            If TypeOf control Is Window Then
                Dim canvas = Form.GetCanvas(control)
                brush = TryCast(canvas.Background, SolidColorBrush)

            ElseIf TypeOf control Is Wpf.Label Then
                Dim L = CType(control, Wpf.Label)
                If TypeOf L.Content Is System.Windows.Shapes.Shape Then
                    brush = CType(L.Content, System.Windows.Shapes.Shape).Fill
                Else
                    brush = TryCast(control.Background, SolidColorBrush)
                End If

            Else
                brush = TryCast(control.Background, SolidColorBrush)
            End If

            If brush IsNot Nothing Then
                Return brush.Color
            Else
                Return control.GetValue(BackColorProperty)
            End If
        End Function

        <ExProperty>
        Public Shared Sub SetBackColor(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim _color = If(
                            Color.IsNone(value),
                            CType(Nothing, Media.Color?),
                            Color.FromString(value)
                        )
                        Dim obj = GetControl(controlName)
                        ' Remove any animation effect to allow setting the new value
                        obj.BeginAnimation(BackColorProperty, Nothing)
                        obj.SetValue(BackColorProperty, _color)
                        ' If TypeOf obj Is Wpf.ComboBox Then ComboBox.SetBackColor(obj, _color)

                    Catch ex As Exception
                        ReportPropertyError(controlName, "BackColor", value, ex)
                    End Try
                End Sub)
        End Sub


        Private Shared ReadOnly ForeColorProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("ForeColor",
                           GetType(Media.Color?), GetType(Wpf.Control))

        ''' <summary>
        ''' The foregeound color used to draw the text of the control.
        ''' Use values from the Colors object, such as Colors.Red
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
                              GetForeColor = New Primitive(Color.GetHexaName(brush))
                          Else
                              GetForeColor = New Primitive(Color.GetHexaName(CType(c.GetValue(ForeColorProperty), Media.Color?)))
                          End If

                      Catch ex As Exception
                          ReportError(controlName, "ForeColor", ex)
                      End Try
                  End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetForeColor(controlName As Primitive, value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Dim c = GetControl(controlName)
                           If Color.IsNone(value) Then
                               c.Foreground = Nothing
                               c.SetValue(ForeColorProperty, Nothing)
                           Else
                               Dim _color = Color.FromString(value)
                               c.Foreground = New SolidColorBrush(_color)
                               c.SetValue(ForeColorProperty, _color)
                           End If

                       Catch ex As Exception
                           ReportPropertyError(controlName, "ForeColor", value, ex)
                       End Try
                   End Sub)
        End Sub

        ''' <summary>
        ''' Shows the color dialog to allow user to change the back color of the current control
        ''' </summary>
        ''' <returns>The color that the user choosed, or "" if he cancelled the operation</returns>
        <ReturnValueType(VariableType.Color)>
        <ExMethod>
        Public Shared Function ChooseBackColor(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Dim color = WinForms.Color.ShowDialog(GetBackColor(controlName))

                    If Not color.IsEmpty Then
                        SetBackColor(controlName, color)
                    End If
                    ChooseBackColor = color
                End Sub)
        End Function


        ''' <summary>
        ''' Shows the color dialog to allow user to change the fore color of the current control
        ''' </summary>
        ''' <returns>The color that the user choosed, or "" if he cancelled the operation</returns>
        <ReturnValueType(VariableType.Color)>
        <ExMethod>
        Public Shared Function ChooseForeColor(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Dim color = WinForms.Color.ShowDialog(GetForeColor(controlName))
                    If Not color.IsEmpty Then
                        SetForeColor(controlName, color)
                    End If
                    ChooseForeColor = color
                End Sub)
        End Function


        ''' <summary>
        ''' Shows the font dialog to allow user to change the font properties including the fore color of the current control
        ''' </summary>
        ''' <returns>The font properties that the user choosed, or "" if he cancelled the operation</returns>
        <ReturnValueType(VariableType.Array)>
        <ExMethod>
        Public Shared Function ChooseFont(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Dim font = Desktop.ShowFontDialog(GetFont(controlName))
                    If Not font.IsEmpty Then
                        SetFont(controlName, font)
                    End If
                    ChooseFont = font
                End Sub)
        End Function

        ''' <summary>
        ''' Gets or sets an array containing the font properties under the keys: Name, Size, Bold, Italic, Underlined and Color, so you can use them as dynamic properties.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Array)>
        <ExProperty>
        Public Shared Function GetFont(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim font As New Primitive
                        Dim c = GetControl(controlName)
                        Dim key As New Primitive("Name")
                        font.Items(key) = New Primitive(c.FontFamily.Source)

                        key._stringValue = "Size"
                        font.Items(key) = System.Math.Round(c.FontSize * 0.75, 1, MidpointRounding.AwayFromZero)

                        key._stringValue = "Bold"
                        font.Items(key) = (c.FontWeight = FontWeights.Bold)

                        key._stringValue = "Italic"
                        font.Items(key) = (c.FontStyle = FontStyles.Italic)

                        Dim cc = TryCast(c, Wpf.ContentControl)
                        key._stringValue = "Underlined"
                        If cc Is Nothing Then
                            Dim txt = TryCast(c, Wpf.TextBox)
                            If txt Is Nothing Then
                                font.Items(key) = False
                            Else
                                font.Items(key) = txt.TextDecorations Is TextDecorations.Underline
                            End If
                        Else
                            Dim tb = TryCast(cc.Content, Wpf.TextBlock)
                            If tb Is Nothing Then
                                font.Items(key) = False
                            Else
                                font.Items(key) = tb.TextDecorations Is TextDecorations.Underline
                            End If
                        End If

                        key._stringValue = "Color"
                        font.Items(key) = GetForeColor(controlName)

                        GetFont = font
                    Catch ex As Exception
                        ReportError(controlName, "Font", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetFont(controlName As Primitive, font As Primitive)
            If font.IsEmpty Then Return

            App.Invoke(
                Sub()
                    Try
                        Dim c = GetControl(controlName)
                        Dim key As New Primitive("Name")
                        If font.ContainsKey(key) Then
                            c.FontFamily = New FontFamily(font.Items(key))
                        End If

                        key._stringValue = "Size"
                        If font.ContainsKey(key) Then
                            c.FontSize = font.Items(key) * 4 / 3
                        End If

                        key._stringValue = "Bold"
                        If font.ContainsKey(key) Then
                            c.FontWeight = If(font.Items(key), FontWeights.Bold, FontWeights.Normal)
                        End If

                        key._stringValue = "Italic"
                        If font.ContainsKey(key) Then
                            c.FontStyle = If(font.Items(key), FontStyles.Italic, FontStyles.Normal)
                        End If

                        key._stringValue = "Color"
                        If font.ContainsKey(key) Then
                            SetForeColor(controlName, font.Items(key))
                        End If

                        key._stringValue = "Underlined"
                        If Not CBool(font.ContainsKey(key)) Then Return

                        Dim cc = TryCast(c, Wpf.ContentControl)
                        If cc Is Nothing Then
                            Dim txt = TryCast(c, Wpf.TextBox)
                            If txt IsNot Nothing Then
                                txt.TextDecorations = If(font.Items(key), TextDecorations.Underline, Nothing)
                            End If
                        Else
                            Dim tb = TryCast(cc.Content, Wpf.TextBlock)
                            If tb IsNot Nothing Then
                                tb.TextDecorations = If(font.Items(key), TextDecorations.Underline, Nothing)
                            End If
                        End If

                    Catch ex As Exception
                        ReportPropertyError(controlName, "FontName", font, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Gets or sets the font name used to display the text of the control
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetFontName(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim _fontFamily = GetControl(controlName).FontFamily
                        GetFontName = New Primitive(If(_fontFamily IsNot Nothing, _fontFamily.Source, "Tahoma"))

                    Catch ex As Exception
                        ReportError(controlName, "FontName", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetFontName(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetControl(controlName).FontFamily = New FontFamily(value)
                    Catch ex As Exception
                        ReportPropertyError(controlName, "FontName", value, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Gets or sets the font size used to display the text of the control.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetFontSize(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetFontSize = System.Math.Round(
                                GetControl(controlName).FontSize * 0.75,
                                1, MidpointRounding.AwayFromZero)

                    Catch ex As Exception
                        ReportError(controlName, "FontSize", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetFontSize(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim size = System.Math.Max(1, value.AsDecimal)
                        GetControl(controlName).FontSize = size * 4 / 3
                    Catch ex As Exception
                        ReportPropertyError(controlName, "FontSize", value, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Gets or sets whether or not the font used to display the text of the control, is bold.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetFontBold(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetFontBold = (GetControl(controlName).FontWeight = FontWeights.Bold)
                    Catch ex As Exception
                        ReportError(controlName, "FontBold", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetFontBold(controlName As Primitive, value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           GetControl(controlName).FontWeight = If(CBool(value), FontWeights.Bold, FontWeights.Normal)
                       Catch ex As Exception
                           ReportPropertyError(controlName, "FontBold", value, ex)
                       End Try
                   End Sub)
        End Sub


        ''' <summary>
        ''' Gets or sets whether or not the font used to display the text of the control, is italic.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetFontItalic(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetFontItalic = (GetControl(controlName).FontStyle = FontStyles.Italic)
                    Catch ex As Exception
                        ReportError(controlName, "FontItalic", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetFontItalic(controlName As Primitive, value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           GetControl(controlName).FontStyle = If(CBool(value), FontStyles.Italic, FontStyles.Normal)
                       Catch ex As Exception
                           ReportPropertyError(controlName, "FontItalic", value, ex)
                       End Try
                   End Sub)
        End Sub

#End Region

#Region "Angle and rotaion"

        ''' <summary>
        ''' Gets or sets the x-coordinate of the point that the control will rotate arround. By default, the control is rotated around its center.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetRotationCenterX(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
                        Dim x = obj.RenderTransformOrigin.X * obj.ActualWidth + Canvas.GetLeft(obj)
                        GetRotationCenterX = System.Math.Round(x, 2, MidpointRounding.AwayFromZero)
                    Catch ex As Exception
                        ReportError(controlName, "RotationCenterX", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetRotationCenterX(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
                        Dim x = (CDbl(value) - Canvas.GetLeft(obj)) / obj.ActualWidth
                        obj.RenderTransformOrigin = New Point(x, obj.RenderTransformOrigin.Y)

                    Catch ex As Exception
                        ReportPropertyError(controlName, "RotationCenterX", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the y-coordinate of the point that the control will rotate arround. By default, the control is rotated around its center.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetRotationCenterY(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
                        Dim y = obj.RenderTransformOrigin.Y * obj.ActualHeight + Canvas.GetTop(obj)
                        GetRotationCenterY = System.Math.Round(y, 2, MidpointRounding.AwayFromZero)
                    Catch ex As Exception
                        ReportError(controlName, "RotationCenterY", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetRotationCenterY(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim obj = GetControl(controlName)
                        Dim y = (CDbl(value) - Canvas.GetTop(obj)) / obj.ActualHeight
                        obj.RenderTransformOrigin = New Point(obj.RenderTransformOrigin.X, y)

                    Catch ex As Exception
                        ReportPropertyError(controlName, "RotationCenterY", value, ex)
                    End Try
                End Sub)
        End Sub

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
                        ReportError(controlName, "Angle", ex)
                    End Try
                End Sub)
        End Function

        Private Shared Function GetAngle(element As DependencyObject) As Double
            If element Is Nothing Then
                Throw New ArgumentNullException("element")
            End If

            Dim angle As Double = element.GetValue(AngleProperty)
            If System.Math.Abs(angle) >= 360 Then
                angle = angle Mod 360
            End If
            Return angle
        End Function


        <ExProperty>
        Public Shared Sub SetAngle(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        SetAngle(GetControl(controlName), value)
                    Catch ex As Exception
                        ReportPropertyError(controlName, "Angle", value, ex)
                    End Try
                End Sub)
        End Sub

        Friend Shared Sub SetAngle(element As DependencyObject, value As Double)
            If element Is Nothing Then
                Throw New ArgumentNullException("element")
            End If

            CType(element, Wpf.Control).BeginAnimation(AngleProperty, Nothing)
            element.SetValue(AngleProperty, value)
        End Sub


        Public Shared ReadOnly AngleProperty As _
                               DependencyProperty = DependencyProperty.RegisterAttached("Angle",
                               GetType(Double), GetType(Wpf.Control),
                               New PropertyMetadata(0.0, AddressOf AngleChanged))

        Private Shared Sub AngleChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim obj = CType(d, Wpf.Control)
            App.Invoke(
                 Sub()
                     Dim rotation = TryCast(obj.RenderTransform, RotateTransform)
                     If rotation Is Nothing Then
                         rotation = New RotateTransform
                         obj.RenderTransform = rotation
                     End If
                     rotation.Angle = CDbl(e.NewValue)
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
                App.Invoke(Sub() SetAngle(obj, GetAngle(obj) + CDbl(angle)))
            Catch ex As Exception
                ReportSubError(controlName, "Rotate", ex)
            End Try
        End Sub

#End Region

#Region "Animation"

        ''' <summary>
        ''' Animates the control's backColor to a new color.
        ''' </summary>
        ''' <param name="newColor">
        ''' The new color to animate the control's backColor to.
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
                        If Color.IsNone(newColor) Then newColor = Colors.Transparent
                        Dim animation As New Animation.ColorAnimation() With {
                            .From = GetBackColor(obj),
                            .To = Color.FromString(newColor),
                            .Duration = TimeSpan.FromMilliseconds(duration)
                        }
                        obj.BeginAnimation(BackColorProperty, animation)
                    End Sub)
            Catch ex As Exception
                ReportSubError(controlName, "AnimateColor", ex)
            End Try
        End Sub

        ''' <summary>
        ''' Animates the control's backColor to a new transparency.
        ''' </summary>
        ''' <param name="transparency">The new transparency to animate the backColor to. Use a value between 0 and 100.</param>
        ''' <param name="duration">The time for the animation, in milliseconds.</param>
        <ExMethod>
        Public Shared Sub AnimateTransparency(controlName As Primitive, transparency As Primitive, duration As Primitive)
            Try
                Dim c = GetBackColor(controlName)
                c = Color.ChangeTransparency(c, transparency)
                AnimateColor(controlName, c, duration)
            Catch ex As Exception
                ReportSubError(controlName, "AnimateColor", ex)
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
                ReportSubError(controlName, "AnimatePos", ex)
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
                        GraphicsWindow.DoubleAnimateProperty(obj, FrameworkElement.WidthProperty, CDbl(width), CDbl(duration))
                        GraphicsWindow.DoubleAnimateProperty(obj, FrameworkElement.HeightProperty, CDbl(height), CDbl(duration))
                    End Sub)
            Catch ex As Exception
                ReportSubError(controlName, "AnimateSize", ex)
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
                        ' Normalize the angle
                        Dim startAngle As Double = obj.GetValue(AngleProperty)
                        If System.Math.Abs(startAngle) >= 360 Then
                            startAngle = startAngle Mod 360
                            SetAngle(obj, startAngle)
                        End If

                        GraphicsWindow.DoubleAnimateProperty(
                            obj,
                            AngleProperty,
                            CDbl(angle),
                            CDbl(duration)
                        )
                    End Sub)

            Catch ex As Exception
                ReportSubError(controlName, "AnimateAngle", ex)
            End Try
        End Sub

#End Region

#Region "Events"

        Friend Shared RemoveEventHandlerActions As New Dictionary(Of String, Action)

        Friend Shared senderAssembly As String
        Friend Shared Sub RemovePrevEventHandler(controlName As String, eventName As String, [removeHandler] As Action)
            Dim key = (senderAssembly & ":" & controlName & "." & eventName).ToLower()
            If RemoveEventHandlerActions.ContainsKey(key) Then
                RemoveEventHandlerActions(key)?.Invoke()
            End If

            RemoveEventHandlerActions(key) = [removeHandler]
        End Sub

        <HideFromIntellisense>
        Public Shared Sub HandleEvents(controlName As Primitive)
            senderAssembly = Reflection.Assembly.GetCallingAssembly().GetName().Name
            [Event]._senderControl = controlName
        End Sub

        Shared Function GetSender(eventNmae As String, Optional getForm As Boolean = False) As FrameworkElement
            Try
                Dim _sender = CType(GetControl([Event].SenderControl), FrameworkElement)
                If TypeOf _sender Is Window AndAlso Not getForm Then
                    App.Invoke(
                           Sub()
                               Dim win = CType(_sender, Window)
                               Dim canvas = Form.GetCanvas(win)
                               If canvas.Background IsNot Nothing OrElse win.AllowsTransparency Then
                                   ' Use the camvas events instead of the form, because form events will not fire if the form is transperent or if the canvas has a non-transparent back color
                                   _sender = canvas
                               End If
                           End Sub)
                End If
                Return _sender

            Catch ex As Exception
                [Event].ShowErrorMessage(eventNmae, ex)
            End Try

            Return Nothing
        End Function

        ''' <summary>
        ''' Removes the handler of the given event of the current control.
        ''' </summary>
        ''' <param name="eventName">The name of the event to remove its handler</param>
        <ExMethod>
        Public Shared Sub RemoveEventHandler(controlName As Primitive, eventName As Primitive)
            If LCase(eventName) = "ontick" Then
                WinTimer.RemoveOnTickHandler(controlName)
            Else
                senderAssembly = Reflection.Assembly.GetCallingAssembly().GetName().Name
                RemovePrevEventHandler(controlName, eventName, Nothing)
            End If
        End Sub


        ''' <summary>
        ''' Fired when user presses the left mouse-button down
        ''' </summary>
        Public Shared Custom Event OnMouseLeftDown As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnMouseLeftDown))
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                If TypeOf _sender Is Wpf.Canvas Then
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnMouseLeftDown),
                            Sub() RemoveHandler _sender.MouseLeftButtonDown, h
                     )
                    AddHandler _sender.MouseLeftButtonDown, h

                Else
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnMouseLeftDown),
                            Sub() RemoveHandler _sender.PreviewMouseLeftButtonDown, h
                     )
                    AddHandler _sender.PreviewMouseLeftButtonDown, h
                End If
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when user clicks the control by the left mouse button.
        ''' </summary>
        Public Shared Custom Event OnClick As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnClick))
                Dim btn = TryCast(_sender, Wpf.Primitives.ButtonBase)
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                If btn IsNot Nothing Then
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnClick),
                            Sub() RemoveHandler btn.Click, h
                    )
                    AddHandler btn.Click, h

                Else
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnClick),
                            Sub() RemoveHandler _sender.MouseLeftButtonUp, h
                    )
                    AddHandler _sender.MouseLeftButtonUp, h
                End If
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        '''     ''' Fired when user releases the left mouse-button
        ''' </summary>
        Public Shared Custom Event OnMouseLeftUp As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnMouseLeftUp))
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                If TypeOf _sender Is Wpf.Canvas Then
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnMouseLeftUp),
                            Sub() RemoveHandler _sender.MouseLeftButtonUp, h
                     )
                    AddHandler _sender.MouseLeftButtonUp, h

                Else
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnMouseLeftUp),
                            Sub() RemoveHandler _sender.PreviewMouseLeftButtonUp, h
                     )
                    AddHandler _sender.PreviewMouseLeftButtonUp, h
                End If
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        '''     ''' Fired when user double-clicks the mouse-button
        ''' </summary>
        Public Shared Custom Event OnDoubleClick As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnDoubleClick))
                Dim h = Sub(Sender As Object, e As Input.MouseButtonEventArgs)
                            If e.ClickCount > 1 Then
                                [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                            End If
                        End Sub

                If TypeOf _sender Is Wpf.Canvas Then
                    RemovePrevEventHandler(
                        [Event].SenderControl,
                        NameOf(OnDoubleClick),
                        Sub() RemoveHandler _sender.MouseLeftButtonDown, h
                    )
                    AddHandler _sender.MouseLeftButtonDown, h

                Else
                    RemovePrevEventHandler(
                        [Event].SenderControl,
                        NameOf(OnDoubleClick),
                        Sub() RemoveHandler _sender.PreviewMouseLeftButtonDown, h
                    )
                    AddHandler _sender.PreviewMouseLeftButtonDown, h
                End If
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when user presses the right mouse-button down
        ''' </summary>
        Public Shared Custom Event OnMouseRightDown As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnMouseRightDown))
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                If TypeOf _sender Is Wpf.Canvas Then
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnMouseRightDown),
                            Sub() RemoveHandler _sender.MouseRightButtonDown, h
                    )
                    AddHandler _sender.MouseRightButtonDown, h

                Else
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnMouseRightDown),
                            Sub() RemoveHandler _sender.PreviewMouseRightButtonDown, h
                    )
                    AddHandler _sender.PreviewMouseRightButtonDown, h
                End If
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when user releases the right mouse-button
        ''' </summary>
        Public Shared Custom Event OnMouseRightUp As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnMouseRightUp))
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                If TypeOf _sender Is Wpf.Canvas Then
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnMouseRightUp),
                            Sub() RemoveHandler _sender.MouseRightButtonUp, h
                    )
                    AddHandler _sender.MouseRightButtonUp, h

                Else
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnMouseRightUp),
                            Sub() RemoveHandler _sender.PreviewMouseRightButtonUp, h
                    )
                    AddHandler _sender.PreviewMouseRightButtonUp, h
                End If
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the mouse pointer moves over the control.
        ''' </summary>
        Public Shared Custom Event OnMouseMove As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnMouseMove))
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                RemovePrevEventHandler(
                        [Event].SenderControl,
                        NameOf(OnMouseMove),
                        Sub() RemoveHandler _sender.PreviewMouseMove, h
                )
                AddHandler _sender.PreviewMouseMove, h
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when user moves the mouse wheel
        ''' </summary>
        Public Shared Custom Event OnMouseWheel As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnMouseWheel))
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                If TypeOf _sender Is Wpf.Canvas Then
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnMouseWheel),
                            Sub() RemoveHandler _sender.MouseWheel, h
                    )
                    AddHandler _sender.MouseWheel, h

                Else
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnMouseWheel),
                            Sub() RemoveHandler _sender.PreviewMouseWheel, h
                    )
                    AddHandler _sender.PreviewMouseWheel, h
                End If
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the mouse pointer enters the control area.
        ''' </summary>
        Public Shared Custom Event OnMouseEnter As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnMouseEnter))
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                RemovePrevEventHandler(
                        [Event].SenderControl,
                        NameOf(OnMouseEnter),
                        Sub() RemoveHandler _sender.MouseEnter, h
                )
                AddHandler _sender.MouseEnter, h
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the mouse pointer leaves the control area.
        ''' </summary>
        Public Shared Custom Event OnMouseLeave As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnMouseLeave))
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                RemovePrevEventHandler(
                        [Event].SenderControl,
                        NameOf(OnMouseLeave),
                        Sub() RemoveHandler _sender.MouseLeave, h
                )
                AddHandler _sender.MouseLeave, h
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the user presses a keyboard key down
        ''' </summary>
        Public Shared Custom Event OnKeyDown As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnKeyDown), True)
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                If TypeOf _sender Is Wpf.Canvas Then
                    ' The form should not preview keys, to let controls handle them
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnKeyDown),
                            Sub() RemoveHandler _sender.KeyDown, h
                    )
                    AddHandler _sender.KeyDown, h
                Else
                    ' Preview keydown can habdle space and other keys im TextBox                    
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnKeyDown),
                            Sub() RemoveHandler _sender.PreviewKeyDown, h
                    )
                    AddHandler _sender.PreviewKeyDown, h
                End If
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the user releases a keyboard key.
        ''' </summary>
        Public Shared Custom Event OnKeyUp As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnKeyUp), True)
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                If TypeOf _sender Is Wpf.Canvas Then
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnKeyUp),
                            Sub() RemoveHandler _sender.KeyUp, h
                    )
                    AddHandler _sender.KeyUp, h

                Else
                    RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnKeyUp),
                            Sub() RemoveHandler _sender.PreviewKeyUp, h
                    )
                    AddHandler _sender.PreviewKeyUp, h
                End If
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event


        Friend Shared Function GetHasFocus(ByVal element As DependencyObject) As Boolean
            If element Is Nothing Then
                Throw New ArgumentNullException("element")
            End If

            Return element.GetValue(HasFocusProperty)
        End Function

        Friend Shared Sub SetHasFocus(ByVal element As DependencyObject, ByVal value As Boolean)
            If element Is Nothing Then
                Throw New ArgumentNullException("element")
            End If

            element.SetValue(HasFocusProperty, value)
        End Sub

        Public Shared ReadOnly HasFocusProperty As _
                               DependencyProperty = DependencyProperty.RegisterAttached("HasFocus",
                               GetType(Boolean), GetType(Control),
                               New PropertyMetadata(False))

        Friend Shared Sub OnLostKeyboardFocus(sender As Object, e As Input.KeyboardFocusChangedEventArgs)
            Dim c = CType(sender, Wpf.Control)
            If Not c.IsKeyboardFocusWithin Then Control.SetHasFocus(c, False)
        End Sub

        ''' <summary>
        ''' Fired when the control gets the focus.
        ''' </summary>
        Public Shared Custom Event OnGotFocus As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnGotFocus))
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            If GetHasFocus(Sender) Then Return
                            SetHasFocus(Sender, True)
                            [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                RemovePrevEventHandler(
                        [Event].SenderControl,
                        NameOf(OnGotFocus),
                        Sub() RemoveHandler _sender.GotKeyboardFocus, h
                )
                AddHandler _sender.GotKeyboardFocus, h
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the control looses the focus.
        ''' </summary>
        Public Shared Custom Event OnLostFocus As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnLostFocus))
                Dim h = Sub(Sender As Object, e As RoutedEventArgs)
                            Dim fw = CType(Sender, FrameworkElement)
                            If Not fw.IsKeyboardFocusWithin Then
                                [Event].HandleEvent(fw, e, handler)
                            End If
                        End Sub

                RemovePrevEventHandler(
                        [Event].SenderControl,
                        NameOf(OnLostFocus),
                        Sub() RemoveHandler _sender.LostFocus, h
                )
                AddHandler _sender.LostFocus, h
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

#End Region

        ''' <summary>
        ''' Brings the current control to front of all other controls on the form.
        ''' </summary>
        <ExMethod>
        Public Shared Sub BringToFront(controlName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim control = GetControl(controlName)
                        If TypeOf control Is Window Then
                            CType(control, Window).Activate()
                            Return
                        End If

                        Dim prefix = GetFormName(controlName) & "."
                        Dim zIndex = Wpf.Panel.GetZIndex(control)

                        For Each key In Forms._controls.Keys
                            If key.StartsWith(prefix) Then
                                Dim zId = Wpf.Panel.GetZIndex(GetControl(key))
                                If zId > zIndex Then zIndex = zId
                            End If
                        Next
                        Wpf.Panel.SetZIndex(control, zIndex + 1)

                    Catch ex As Exception
                        ReportSubError(controlName, "BringToFront", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Sends the current control to the back of all other controls on the form.
        ''' </summary>
        <ExMethod>
        Public Shared Sub SendToBack(controlName As Primitive)
            App.Invoke(
                 Sub()
                     Try
                         Dim control = GetControl(controlName)
                         Dim prefix = GetFormName(controlName) & "."
                         Dim zIndex = Wpf.Panel.GetZIndex(control)

                         For Each key In Forms._controls.Keys
                             If key.StartsWith(prefix) Then
                                 Dim zId = Wpf.Panel.GetZIndex(GetControl(key))
                                 If zId < zIndex Then zIndex = zId
                             End If
                         Next
                         Wpf.Panel.SetZIndex(control, zIndex - 1)

                     Catch ex As Exception
                         ReportSubError(controlName, "SendToBack", ex)
                     End Try
                 End Sub)
        End Sub

        Private Shared Sub CloneMenuItemsRecursive(formName As String, sourceItem As Wpf.MenuItem, targetItem As Wpf.MenuItem)
            targetItem.Header = sourceItem.Header
            targetItem.IsCheckable = sourceItem.IsCheckable
            targetItem.IsChecked = sourceItem.IsChecked
            targetItem.InputGestureText = sourceItem.InputGestureText

            Dim menuItemName = formName & "." & sourceItem.Name

            Dim key = (senderAssembly & ":" & menuItemName & ".OnClick").ToLower()
            If MenuItem.MenuEventHandlers.ContainsKey(key) Then
                AddHandler targetItem.Click, MenuItem.MenuEventHandlers(key)
            End If

            If sourceItem.IsCheckable Then
                key = (senderAssembly & ":" & menuItemName & ".OnCheck").ToLower()
                If MenuItem.MenuEventHandlers.ContainsKey(key) Then
                    Dim h = MenuItem.MenuEventHandlers(key)
                    AddHandler targetItem.Checked, h
                    AddHandler targetItem.Unchecked, h
                End If
            End If

            If sourceItem.HasItems Then
                For Each subItem As Wpf.Control In sourceItem.Items
                    If TypeOf subItem Is Wpf.Separator Then
                        targetItem.Items.Add(New Wpf.Separator())
                    Else
                        Dim clonedSubItem As New Wpf.MenuItem()
                        CloneMenuItemsRecursive(formName, subItem, clonedSubItem)
                        targetItem.Items.Add(clonedSubItem)
                    End If
                Next

                key = (senderAssembly & ":" & sourceItem.Name & ".OnOpen").ToLower()
                If MenuItem.MenuEventHandlers.ContainsKey(key) Then
                    AddHandler targetItem.SubmenuOpened, MenuItem.MenuEventHandlers(key)
                End If
            End If
        End Sub

        ''' <summary>
        ''' Sets the context menu of the control.
        ''' </summary>
        ''' <param name="sourceMenuName">The menu you want to show its items in the context menu.</param>
        ''' <param name="clone">
        ''' Use True to copy the items from the source menu. They will have the same properties and event handlers as the original items. This is helpful when you need to show the same commands in the main menu and the context menu, like the edit command of the textbox.
        ''' Use False to remove the  items from the source menu and add them to the context. In such case, you should hide the source menu from the user. This is helpful when you need a unique context menu, so you can create its items using the menu designer then use this method to move them to the context menu.
        ''' </param>
        <ExMethod>
        Public Shared Sub SetContextMenu(controlName As Primitive, sourceMenuName As Primitive, clone As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim control = GetControl(controlName)
                        Dim mi = MenuItem.GetMenuItem(sourceMenuName)
                        Dim contextMenu As New ContextMenu()
                        control.ContextMenu = contextMenu

                        If Not CBool(clone) Then
                            Dim RemovedItems As New List(Of Wpf.Control)
                            For Each item In mi.Items
                                RemovedItems.Add(item)
                            Next

                            mi.Items.Clear()

                            For Each item In RemovedItems
                                contextMenu.Items.Add(item)
                            Next
                            Return
                        End If

                        Dim key = (senderAssembly & ":" & sourceMenuName.AsString() & ".OnOpen").ToLower()
                        If MenuItem.MenuEventHandlers.ContainsKey(key) Then
                            Dim h = MenuItem.MenuEventHandlers(key)
                            Dim SourceItems = mi.Items
                            AddHandler contextMenu.Opened,
                                Sub(Sender As Wpf.Control, e As RoutedEventArgs)
                                    h.Invoke(Sender, e)
                                    Dim items = contextMenu.Items
                                    For i = 0 To items.Count - 1
                                        Dim menuItem = TryCast(items(i), Wpf.MenuItem)
                                        If menuItem IsNot Nothing Then
                                            Dim sourceItem = CType(SourceItems(i), Wpf.MenuItem)
                                            menuItem.Visibility = sourceItem.Visibility
                                            menuItem.IsEnabled = sourceItem.IsEnabled
                                            menuItem.IsChecked = sourceItem.IsChecked
                                        End If
                                    Next
                                End Sub
                        End If

                        For Each item As Wpf.Control In mi.Items
                            If TypeOf item Is Wpf.Separator Then
                                contextMenu.Items.Add(New Wpf.Separator())
                            Else
                                Dim clonedItem As New Wpf.MenuItem()
                                CloneMenuItemsRecursive(GetFormName(controlName), item, clonedItem)
                                contextMenu.Items.Add(clonedItem)
                            End If
                        Next

                    Catch ex As Exception
                        ReportSubError(controlName, "SetContextMenu", ex)
                    End Try
                End Sub)
        End Sub

        Private Shared Function GetFormName(controlName As Primitive) As String
            Dim name = controlName.AsString().ToLower()
            Dim dotPos = name.IndexOf(".")
            If dotPos = -1 Then Return name
            Return name.Substring(0, dotPos)
        End Function

        Friend Shared Function GetControl(key As String) As Wpf.Control
            Return GetFrameworkElement(key)
        End Function

        Friend Shared Function GetFrameworkElement(key As String) As FrameworkElement
            key = key.ToLower()
            If Not key.Contains(".") Then Return Forms._forms(key)

            If Not Forms._controls.ContainsKey(key) Then
                Dim names = key.Split("."c)
                If names(0) = names(1) Then Return Forms._forms(names(0))

                Throw New Exception($"There is no control named `{names(1)}` on form {names(0)}.")
            End If

            Return Forms._controls(key)
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