
Imports System.Windows
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

        Shared Sub ReportError(formName As String, controlName As String, memberName As String, ex As Exception)
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
                        GetName = GetControl(controlName).Name
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
                                GetParentForm = name
                            Else
                                GetParentForm = ""
                            End If
                        Else
                            GetParentForm = name.Substring(0, i)
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
                        GetTypeName = GetControl(controlName).GetType().Name
                    Catch ex As Exception
                        ReportError(controlName, "TypeName", ex)
                    End Try
                End Sub)
        End Function


        ''' <summary>
        ''' An extra property to store any value you want that is related to the control. you can store multiple values as by putting them in an array or a dynamic object
        ''' </summary>
        <ExProperty>
        Public Shared Function GetTag(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetTag = CStr(GetControl(controlName).Tag)
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
        ''' Gets or sets the x-pos of the control on its parent control.
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
        ''' Gets or sets the y-pos of the control on its parent control.
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
                        Wpf.Canvas.SetTop(obj, value)
                    Catch ex As Exception
                        ReportPropertyError(controlName, "Top", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the width of the control.
        ''' If you want an auto-width to fit the control's content width, set this property to -1. 
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
                        If w < 0 Then w = Double.NaN
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
        ''' If you want an auto-heigh to fit the control's content heigh, set this property to -1. 
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
                        If h < 0 Then h = Double.NaN
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
        ''' When it is True, the control is shown at the form.
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
            App.Invoke(
                Sub()
                    Try
                        GetControl(controlName).Visibility = If(value, Visibility.Visible, Visibility.Hidden)
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
                           GetControl(controlName).FlowDirection = If(CBool(value), FlowDirection.RightToLeft, FlowDirection.LeftToRight)
                       Catch ex As Exception
                           ReportPropertyError(controlName, "RightToLeft", value, ex)
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
                        GetMouseX = System.Math.Round(Input.Mouse.GetPosition(c).X, MidpointRounding.AwayFromZero)
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
                        GetMouseY = System.Math.Round(Input.Mouse.GetPosition(c).Y, MidpointRounding.AwayFromZero)
                    Catch ex As Exception
                        ReportError(controlName, "MouseY", ex)
                    End Try
                End Sub)
        End Function

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
                        GetToolTip = If(c.GetValue(TipProperty), "").ToString()
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
                         GetBackColor = Color.GetHexaName(GetBackColor(c))
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
                              GetForeColor = Color.GetHexaName(brush)
                          Else
                              GetForeColor = Color.GetHexaName(CType(c.GetValue(ForeColorProperty), Media.Color?))
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
            Dim color = WinForms.Color.ShowDialog(GetBackColor(controlName))
            If Not color.IsEmpty Then
                SetBackColor(controlName, color)
            End If
            Return color
        End Function


        ''' <summary>
        ''' Shows the color dialog to allow user to change the fore color of the current control
        ''' </summary>
        ''' <returns>The color that the user choosed, or "" if he cancelled the operation</returns>
        <ReturnValueType(VariableType.Color)>
        <ExMethod>
        Public Shared Function ChooseForeColor(controlName As Primitive) As Primitive
            Dim color = WinForms.Color.ShowDialog(GetForeColor(controlName))
            If Not color.IsEmpty Then
                SetForeColor(controlName, color)
            End If
            Return color
        End Function


        ''' <summary>
        ''' Shows the font dialog to allow user to change the font properties including the fore color of the current control
        ''' </summary>
        ''' <returns>The font properties that the user choosed, or "" if he cancelled the operation</returns>
        <ReturnValueType(VariableType.Array)>
        <ExMethod>
        Public Shared Function ChooseFont(controlName As Primitive) As Primitive
            Dim font = Desktop.ShowFontDialog(GetFont(controlName))
            If Not font.IsEmpty Then
                SetFont(controlName, font)
            End If
            Return font
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
                        font.Items("Name") = c.FontFamily.Source
                        font.Items("Size") = System.Math.Round(c.FontSize * 0.75, 1, MidpointRounding.AwayFromZero)
                        font.Items("Bold") = (c.FontWeight = FontWeights.Bold)
                        font.Items("Italic") = (c.FontStyle = FontStyles.Italic)

                        Dim cc = TryCast(c, Wpf.ContentControl)
                        If cc Is Nothing Then
                            Dim txt = TryCast(c, Wpf.TextBox)
                            If txt Is Nothing Then
                                font.Items("Underlined") = False
                            Else
                                font.Items("Underlined") = txt.TextDecorations Is TextDecorations.Underline
                            End If
                        Else
                            Dim tb = TryCast(cc.Content, Wpf.TextBlock)
                            If tb Is Nothing Then
                                font.Items("Underlined") = False
                            Else
                                font.Items("Underlined") = tb.TextDecorations Is TextDecorations.Underline
                            End If
                        End If

                        font.Items("Color") = GetForeColor(controlName)

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
                        If font.ContainsKey("Name") Then
                            c.FontFamily = New FontFamily(font.Items("Name"))
                        End If

                        If font.ContainsKey("Size") Then
                            c.FontSize = font.Items("Size") * 4 / 3
                        End If

                        If font.ContainsKey("Bold") Then
                            c.FontWeight = If(font.Items("Bold"), FontWeights.Bold, FontWeights.Normal)
                        End If

                        If font.ContainsKey("Italic") Then
                            c.FontStyle = If(font.Items("Italic"), FontStyles.Italic, FontStyles.Normal)
                        End If

                        If font.ContainsKey("Color") Then
                            SetForeColor(controlName, font.Items("Color"))
                        End If

                        If Not CBool(font.ContainsKey("Underlined")) Then Return

                        Dim cc = TryCast(c, Wpf.ContentControl)
                        If cc Is Nothing Then
                            Dim txt = TryCast(c, Wpf.TextBox)
                            If txt IsNot Nothing Then
                                txt.TextDecorations = If(font.Items("Underlined"), TextDecorations.Underline, Nothing)
                            End If
                        Else
                            Dim tb = TryCast(cc.Content, Wpf.TextBlock)
                            If tb IsNot Nothing Then
                                tb.TextDecorations = If(font.Items("Underlined"), TextDecorations.Underline, Nothing)
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
                        GetFontName = If(_fontFamily IsNot Nothing, _fontFamily.Source, "Tahoma")

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
                        GetFontSize = System.Math.Round(GetControl(controlName).FontSize * 0.75, 1, MidpointRounding.AwayFromZero)
                    Catch ex As Exception
                        ReportError(controlName, "FontSize", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetFontSize(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetControl(controlName).FontSize = value * 4 / 3
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

            Return NormalizeAngle(element.GetValue(AngleProperty))
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
            element.SetValue(AngleProperty, NormalizeAngle(value))
        End Sub

        Private Shared Function NormalizeAngle(value As Double) As Double
            If value >= 360 Then
                Return value Mod 360
            ElseIf value < 0 Then
                Dim v = (value Mod 360)
                Return If(v < 0, 360 + v, 0)
            Else
                Return value
            End If
        End Function

        Public Shared ReadOnly AngleProperty As _
                               DependencyProperty = DependencyProperty.RegisterAttached("Angle",
                               GetType(Double), GetType(Wpf.Control),
                               New PropertyMetadata(0.0, AddressOf AngleChanged))

        Private Shared Sub AngleChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim obj = CType(d, Wpf.Control)
            App.BeginInvoke(
                 Sub()
                     Dim rotation = TryCast(obj.RenderTransform, RotateTransform)
                     If rotation Is Nothing Then
                         rotation = New RotateTransform
                         rotation.CenterX = obj.ActualWidth / 2.0
                         rotation.CenterY = obj.ActualHeight / 2.0
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
                App.BeginInvoke(Sub() SetAngle(obj, GetAngle(obj) + CDbl(angle)))
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
        ''' <param name="transparency">
        ''' The new transparency to animate the backColor to. Use a value between 0 and 100.
        ''' </param>
        ''' <param name="duration">
        ''' The time for the animation, in milliseconds.
        ''' </param>
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
                        GraphicsWindow.DoubleAnimateProperty(obj, FrameworkElement.WidthProperty, width, duration)
                        GraphicsWindow.DoubleAnimateProperty(obj, FrameworkElement.HeightProperty, height, duration)
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

        <HideFromIntellisense>
        Public Shared Sub HandleEvents(ControlName As Primitive)
            [Event].SenderControl = ControlName
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
        ''' Fired when user presses the left mouse-button down
        ''' </summary>
        Public Shared Custom Event OnMouseLeftDown As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnMouseLeftDown))
                AddHandler _sender.PreviewMouseLeftButtonDown, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when user releases the left mouse-button
        ''' </summary>
        Public Shared Custom Event OnClick As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnClick))
                AddHandler _sender.PreviewMouseLeftButtonUp,
                    Sub(Sender As Object, e As RoutedEventArgs)
                        [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                    End Sub
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
                AddHandler _sender.PreviewMouseLeftButtonUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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

                If TypeOf _sender Is Wpf.Canvas Then
                    AddHandler _sender.MouseLeftButtonDown,
                        Sub(Sender As Object, e As Input.MouseButtonEventArgs)
                            If e.ClickCount > 1 Then [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                        End Sub

                Else
                    AddHandler _sender.PreviewMouseLeftButtonDown,
                         Sub(Sender As Object, e As Input.MouseButtonEventArgs)
                             If e.ClickCount > 1 Then [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                         End Sub
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
                AddHandler _sender.PreviewMouseRightButtonDown, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                AddHandler _sender.PreviewMouseRightButtonUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                AddHandler _sender.PreviewMouseMove, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                AddHandler _sender.PreviewMouseWheel, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                AddHandler _sender.MouseEnter, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                AddHandler _sender.MouseLeave, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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
                If TypeOf _sender Is Window OrElse TypeOf _sender Is Wpf.Canvas Then
                    ' The form should not preview keys, to let controls handle them
                    AddHandler _sender.KeyDown,
                        Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Else
                    ' Preview keydown can habdle space and other keys im TextBox
                    AddHandler _sender.PreviewKeyDown,
                        Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                End If
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the user presses a keyboard key down on a control or any of its cjild controls.
        ''' </summary>
        Public Shared Custom Event OnPreviewKeyDown As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnPreviewKeyDown), True)
                AddHandler _sender.PreviewKeyDown,
                    Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler, True)
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
                If TypeOf _sender Is Window OrElse TypeOf _sender Is Wpf.Canvas Then
                    AddHandler _sender.KeyUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Else
                    AddHandler _sender.PreviewKeyUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                End If
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the user releases a keyboard key on the control or any of its child controls.
        ''' </summary>
        Public Shared Custom Event OnPreviewKeyUp As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnPreviewKeyUp), True)
                AddHandler _sender.PreviewKeyUp, Sub(Sender As Object, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler, True)
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        Private Shared focusTime As Date = Now

        ''' <summary>
        ''' Fired when the control gets the focus.
        ''' </summary>
        Public Shared Custom Event OnGotFocus As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim _sender = GetSender(NameOf(OnGotFocus))

                AddHandler _sender.GotKeyboardFocus,
                    Sub(Sender As Object, e As RoutedEventArgs)
                        [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                    End Sub
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
                Dim name = [Event].SenderControl.AsString().ToLower()
                AddHandler _sender.LostFocus,
                    Sub(Sender As Object, e As RoutedEventArgs)
                        [Event].EventsHandler(
                                CType(Sender, FrameworkElement),
                                e,
                                handler
                        )
                    End Sub

            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
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
                Return CType(Forms._forms(key.ToLower), Wpf.Control)
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