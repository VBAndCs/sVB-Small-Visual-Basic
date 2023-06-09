Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls

Namespace Library
    ''' <summary>
    ''' The Controls object allows you to add, move and interact with controls.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Controls
        Private Shared _positionMap As New Dictionary(Of String, Point)
        Private Shared _lastClickedButton As Primitive
        Private Shared _lastTypedTextBox As Primitive
        Friend Const GW_NAME As String = "graphicswindow"

        ''' <summary>
        ''' Gets the last Button that was clicked on the Graphics Window.
        ''' </summary>
        Public Shared ReadOnly Property LastClickedButton As Primitive
            Get
                Return _lastClickedButton
            End Get
        End Property

        ''' <summary>
        ''' Gets the last TextBox, text was typed into.
        ''' </summary>
        Public Shared ReadOnly Property LastTypedTextBox As Primitive
            Get
                Return _lastTypedTextBox
            End Get
        End Property

        Private Shared Events As New EventHandlerList


        ''' <summary>
        ''' Raises an event when any button control is clicked.
        ''' </summary>
        Public Shared Custom Event ButtonClicked As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                Dim h = TryCast(Events("ButtonClicked"), SmallVisualBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler("ButtonClicked", h)
                Events.AddHandler("ButtonClicked", Value)
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                Events.RemoveHandler("ButtonClicked", Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events("ButtonClicked"), SmallVisualBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Raises an event when text is typed into any TextBox control.
        ''' </summary>
        Public Shared Custom Event TextTyped As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                Dim h = TryCast(Events("TextTyped"), SmallVisualBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler("TextTyped", h)
                Events.AddHandler("TextTyped", Value)
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                Events.RemoveHandler("TextTyped", Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events("TextTyped"), SmallVisualBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Adds a button to the graphics window at the specified position.
        ''' </summary>
        ''' <param name="caption">The caption to display in the button.</param>
        ''' <param name="left">The x co-ordinate of the button.</param>
        ''' <param name="top">The y co-ordinate of the button.</param>
        ''' <returns>
        ''' The key of the Button. sVB can deal with this key as an object of type Button, so, you can access the Button methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Button)>
        Public Shared Function AddButton(caption As Primitive, left As Primitive, top As Primitive) As Primitive
            Dim name As String = Shapes.GenerateNewName("Button", True)
            GraphicsWindow.Invoke(
                Sub()
                    Dim button1 As New Button With {
                          .Content = caption,
                          .Padding = New Thickness(4.0)
                    }
                    Canvas.SetLeft(button1, left)
                    Canvas.SetTop(button1, top)
                    AddHandler button1.Click, AddressOf OnButtonClicked
                    GraphicsWindow.AddControl(name, button1)
                End Sub)
            Return name
        End Function

        ''' <summary>
        ''' Gets the current caption of the specified button.
        ''' </summary>
        ''' <param name="buttonName">The Button whose caption is requested.</param>
        ''' <returns>
        ''' The current caption of the button.
        ''' </returns>
        Public Shared Function GetButtonCaption(buttonName As Primitive) As Primitive
            Return GraphicsWindow.InvokeWithReturn(
                    Function() As Primitive
                        Dim value As UIElement = Nothing
                        If Not GraphicsWindow._objectsMap.TryGetValue(buttonName, value) Then
                            Return New Primitive("")
                        End If

                        Dim button1 = TryCast(value, Button)
                        If button1 Is Nothing Then
                            Return New Primitive("")
                        End If
                        Return New Primitive(button1.Content.ToString())
                    End Function)
        End Function

        ''' <summary>
        ''' Sets the caption of the specified button.
        ''' </summary>
        ''' <param name="buttonName">The Button whose caption needs to be set.</param>
        ''' <param name="caption">The new caption for the button.</param>
        Public Shared Sub SetButtonCaption(buttonName As Primitive, caption As Primitive)
            Dim value As UIElement = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(buttonName, value) Then
                Return
            End If

            Dim button1 As Button = TryCast(value, Button)
            If button1 IsNot Nothing Then
                GraphicsWindow.BeginInvoke(Sub() button1.Content = caption)
            End If
        End Sub

        ''' <summary>
        ''' Adds a text input box to the graphics window at the specified position.
        ''' </summary>
        ''' <param name="left">The x co-ordinate of the text box.</param>
        ''' <param name="top">The y co-ordinate of the text box.</param>
        ''' <returns>
        ''' The key of the TextBox. sVB can deal with this key as an object of type TextBox, so, you can access the TextBox methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.TextBox)>
        Public Shared Function AddTextBox(left As Primitive, top As Primitive) As Primitive
            Dim name = Shapes.GenerateNewName("TextBox", True)
            GraphicsWindow.Invoke(
                Sub()
                    Dim textBox1 As New TextBox With {
                        .Width = 160.0,
                        .Padding = New Thickness(2.0)
                    }
                    Canvas.SetLeft(textBox1, left)
                    Canvas.SetTop(textBox1, top)
                    AddHandler textBox1.TextChanged, AddressOf OnTextChanged
                    GraphicsWindow.AddControl(name, textBox1)
                End Sub)
            Return name
        End Function

        ''' <summary>
        ''' Adds a new CheckBox control to the graphics window
        ''' </summary>
        ''' <param name="caption">the text to desply on the CheckBox</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="isChecked">The value to set to the Checked property</param>
        ''' <returns>
        ''' The key of the CheckBox. sVB can deal with this key as an object of type CheckBox, so, you can access the CheckBox methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.CheckBox)>
        Public Shared Function AddCheckBox(
                         caption As Primitive,
                         left As Primitive,
                         top As Primitive,
                         isChecked As Primitive
                   ) As Primitive

            Dim name = Shapes.GenerateNewName("CheckBox", True)
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim key = WinForms.Form.AddCheckBox(GW_NAME, name, left, top, caption, isChecked)
                    GraphicsWindow.AddControl(
                        name,
                        WinForms.Control.GetControl(key),
                        False
                    )
                    AddCheckBox = key
                End Sub)
        End Function


        ''' <summary>
        ''' Adds a new ComboBox control to the graphics window
        ''' </summary>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <returns>
        ''' The key of the comboBox. sVB can deal with this key as an object of type ComboBox, so, you can use the ComboBox methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.ComboBox)>
        Public Shared Function AddComboBox(
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive
                  ) As Primitive

            Dim name = Shapes.GenerateNewName("ComboBox", True)
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim key = WinForms.Form.AddComboBox(
                        GW_NAME, name, left, top, width, height
                    )
                    GraphicsWindow.AddControl(
                        name,
                        WinForms.Control.GetControl(key),
                        False
                    )
                    AddComboBox = key
                End Sub)
        End Function


        ''' <summary>
        ''' Adds a new DatePicker control to the graphics window
        ''' </summary>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="selectedDate">the date that will be selected in the control</param>
        ''' <returns>
        ''' The key of the DatePicker. sVB can deal with this key as an object of type DatePicker, so, you can access the DatePicker methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.DatePicker)>
        Public Shared Function AddDatePicker(
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         selectedDate As Primitive
                    ) As Primitive

            Dim name = Shapes.GenerateNewName("DatePicker", True)
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim key = WinForms.Form.AddDatePicker(
                        GW_NAME, name, left, top, width, selectedDate
                    )
                    GraphicsWindow.AddControl(
                        name,
                        WinForms.Control.GetControl(key),
                        False
                    )
                    AddDatePicker = key
                End Sub)
        End Function

        ''' <summary>
        ''' Adds a new Label control to the graphics window
        ''' </summary>
        ''' <param name="caption">The text to display on the label</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <returns>
        ''' The key of the label. sVB can deal with this key as an object of type Label, so, you can access the Label methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Label)>
        Public Shared Function AddLabel(
                         caption As Primitive,
                         left As Primitive,
                         top As Primitive
                   ) As Primitive

            Dim name = Shapes.GenerateNewName("Label", True)
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim key = WinForms.Form.AddLabel(
                        GW_NAME, name, left, top, -1, -1
                    )
                    GraphicsWindow.AddControl(
                        name,
                        WinForms.Control.GetControl(key),
                        False
                    )
                    WinForms.Label.SetText(key, caption)
                    AddLabel = key
                End Sub)
        End Function


        ''' <summary>
        ''' Adds a new ListBox control to the graphics window
        ''' </summary>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <returns>
        ''' The key of the listBox. sVB can deal with this key as an object of type ListBox, so, you can access the ListBox methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.ListBox)>
        Public Shared Function AddListBox(
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive
                  ) As Primitive

            Dim name = Shapes.GenerateNewName("ListBox", True)
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim key = WinForms.Form.AddListBox(
                        GW_NAME, name, left, top, width, height
                    )
                    GraphicsWindow.AddControl(
                        name,
                        WinForms.Control.GetControl(key),
                        False
                    )
                    AddListBox = key
                End Sub)
        End Function

        ''' <summary>
        ''' Adds a new ProgressBar control to the graphics window
        ''' </summary>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <param name="minimum">The progress minimum value</param>
        ''' <param name="maximum">The progress maximum value. Use 0 if the max value is indeterminate.</param>
        ''' <returns>
        ''' The key of the ProgressBar. sVB can deal with this key as an object of type ProgressBar, so, you can access the ProgressBar methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.ProgressBar)>
        Public Shared Function AddProgressBar(
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive,
                         minimum As Primitive,
                         maximum As Primitive
                   ) As Primitive

            Dim name = Shapes.GenerateNewName("ProgressBar", True)
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim key = WinForms.Form.AddProgressBar(
                        GW_NAME, name,
                        left, top, width, height,
                        minimum, maximum
                    )
                    GraphicsWindow.AddControl(
                        name,
                        WinForms.Control.GetControl(key),
                        False
                    )
                    AddProgressBar = key
                End Sub)
        End Function

        ''' <summary>
        ''' Adds a new RadioButton control to the graphics window
        ''' </summary>
        ''' <param name="caption">The text to desply on the RadioButton</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="groupName">The name of the group to add the button to</param>
        ''' <param name="isChecked">The value to set to the Checked property</param>
        ''' <returns>
        ''' The key of the RadioButton. sVB can deal with this key as an object of type RadioButton, so, you can access the RadioButton methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.RadioButton)>
        Public Shared Function AddRadioButton(
                         caption As Primitive,
                         left As Primitive,
                         top As Primitive,
                         groupName As Primitive,
                         isChecked As Primitive
                   ) As Primitive

            Dim name = Shapes.GenerateNewName("RadioButton", True)
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim key = WinForms.Form.AddRadioButton(
                        GW_NAME, name,
                        left, top,
                        caption, groupName, isChecked
                    )
                    GraphicsWindow.AddControl(
                        name,
                        WinForms.Control.GetControl(key),
                        False
                    )
                    AddRadioButton = key
                End Sub)
        End Function


        ''' <summary>
        ''' Adds a new ScrollBar control to the graphics window
        ''' </summary>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <param name="minimum">The scrollbar minimum value</param>
        ''' <param name="maximum">The scrollbar maximum value.</param>
        ''' <param name="value">The scrollbar current value</param>
        ''' <returns>
        ''' The key of the Scrollbar. sVB can deal with this key as an object of type Scrollbar, so, you can access the Scrollbar methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.ScrollBar)>
        Public Shared Function AddScrollBar(
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive,
                         minimum As Primitive,
                         maximum As Primitive,
                         value As Primitive
                   ) As Primitive

            Dim name = Shapes.GenerateNewName("ScrollBar", True)
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim key = WinForms.Form.AddScrollBar(
                        GW_NAME, name,
                        left, top, width, height,
                        minimum, maximum, value
                    )
                    GraphicsWindow.AddControl(
                        name,
                        WinForms.Control.GetControl(key),
                        False
                    )
                    AddScrollBar = key
                End Sub)
        End Function


        ''' <summary>
        ''' Adds a new Slider control to the graphics window
        ''' </summary>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <param name="minimum">The slider minimum value</param>
        ''' <param name="maximum">The slider maximum value.</param>
        ''' <param name="value">The slider current value</param>
        ''' <param name="tickFrequency">The distance between slide ticks</param>
        ''' <returns>
        ''' The key of the Slider. sVB can deal with this key as an object of type Slider, so, you can access the Slider methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Slider)>
        Public Shared Function AddSlider(
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive,
                         minimum As Primitive,
                         maximum As Primitive,
                         value As Primitive,
                         tickFrequency As Primitive
                    ) As Primitive

            Dim name = Shapes.GenerateNewName("Slider", True)
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim key = WinForms.Form.AddSlider(
                        GW_NAME, name,
                        left, top, width, height,
                        minimum, maximum, value, tickFrequency
                    )
                    GraphicsWindow.AddControl(
                        name,
                        WinForms.Control.GetControl(key),
                        False
                    )
                    AddSlider = key
                End Sub)
        End Function


        ''' <summary>
        ''' Adds a new Timer control to the graphics window. This allows you to add many timers to do different tasks in different intervals.
        ''' </summary>
        ''' <param name="interval">The delay time in milliseconds between ticks</param>
        ''' <returns>
        ''' The key of the Timer. sVB can deal with this key as an object of type WinTimer, so, you can access the WinTimer methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.WinTimer)>
        Public Shared Function AddTimer(
                         interval As Primitive
                   ) As Primitive

            Dim name = Shapes.GenerateNewName("Timer", False)
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim key = WinForms.Form.AddTimer(
                        GW_NAME, name, interval
                    )
                    AddTimer = key
                End Sub)
        End Function


        ''' <summary>
        ''' Adds a new ToggleButton control to the graphics window
        ''' </summary>
        ''' <param name="caption">The text that will b displayed on the control</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <returns>
        ''' The key of the ToggleButton. sVB can deal with this key as an object of type ToggleButton, so, you can access the ToggleButton methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.ToggleButton)>
        Public Shared Function AddToggleButton(
                         caption As Primitive,
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive
                    ) As Primitive

            Dim name = Shapes.GenerateNewName("ToggleButton", True)
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim key = WinForms.Form.AddToggleButton(
                        GW_NAME, name,
                        left, top, width, height
                    )
                    WinForms.ToggleButton.SetText(key, caption)
                    GraphicsWindow.AddControl(
                        name,
                        WinForms.Control.GetControl(key),
                        False
                    )
                    AddToggleButton = key
                End Sub)
        End Function

        ''' <summary>
        ''' Adds a multi-line text input box to the graphics window at the specified position.
        ''' </summary>
        ''' <param name="left">The x co-ordinate of the text box.</param>
        ''' <param name="top">The y co-ordinate of the text box.</param>
        ''' <returns>
        ''' The key of the TextBox. sVB can deal with this key as an object of type TextBox, so, you can access the TextBox methods via it.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.TextBox)>
        Public Shared Function AddMultiLineTextBox(left As Primitive, top As Primitive) As Primitive
            Dim name As String = Shapes.GenerateNewName("TextBox", True)
            GraphicsWindow.Invoke(
                Sub()
                    Dim textBox1 As New TextBox With {
                           .AcceptsReturn = True,
                           .Width = 200.0,
                           .Height = 80.0,
                           .HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                           .VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                           .Padding = New Thickness(2.0)
                    }
                    Canvas.SetLeft(textBox1, left)
                    Canvas.SetTop(textBox1, top)
                    AddHandler textBox1.TextChanged, AddressOf OnTextChanged
                    GraphicsWindow.AddControl(name, textBox1)
                End Sub)
            Return name
        End Function

        ''' <summary>
        ''' Gets the current text of the specified TextBox.
        ''' </summary>
        ''' <param name="textBoxName">
        ''' The TextBox whose text is requested.
        ''' </param>
        ''' <returns>
        ''' The text in the TextBox
        ''' </returns>
        Public Shared Function GetTextBoxText(textBoxName As Primitive) As Primitive
            Return GraphicsWindow.InvokeWithReturn(
                Function() As Primitive
                    Dim value As UIElement = Nothing
                    If Not GraphicsWindow._objectsMap.TryGetValue(textBoxName, value) Then
                        Return New Primitive("")
                    End If

                    Dim textBox1 As TextBox = TryCast(value, TextBox)
                    If textBox1 Is Nothing Then
                        Return New Primitive("")
                    End If

                    Return New Primitive(textBox1.Text)
                End Function)
        End Function

        ''' <summary>
        ''' Sets the text of the specified TextBox.
        ''' </summary>
        ''' <param name="textBoxName">The TextBox whose text needs to be set.</param>
        ''' <param name="text">The new text for the TextBox.</param>
        Public Shared Sub SetTextBoxText(textBoxName As Primitive, text As Primitive)
            Dim value As UIElement = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(textBoxName, value) Then
                Return
            End If
            Dim textBox1 As TextBox = TryCast(value, TextBox)
            If textBox1 IsNot Nothing Then
                GraphicsWindow.BeginInvoke(Sub() textBox1.Text = text)
            End If
        End Sub

        ''' <summary>
        ''' Removes a control from the Graphics Window.
        ''' </summary>
        ''' <param name="controlName">The name of the control that needs to be removed.</param>
        Public Shared Sub Remove(controlName As Primitive)
            controlName = controlName.AsString().ToLower()
            Dim value As UIElement = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(controlName, value) Then
                Dim button1 As Button = TryCast(value, Button)
                If button1 IsNot Nothing Then
                    RemoveHandler button1.Click, AddressOf OnButtonClicked
                End If

                GraphicsWindow.RemoveShape(controlName)
                WinForms.Form.RemoveControl(GW_NAME, controlName)
            End If
        End Sub

        ''' <summary>
        ''' Moves the control with the specified name to a new position.
        ''' </summary>
        ''' <param name="control">The name of the control to move.</param>
        ''' <param name="x">The x co-ordinate of the new position.</param>
        ''' <param name="y">The y co-ordinate of the new position.</param>
        Public Shared Sub Move(control As Primitive, x As Primitive, y As Primitive)
            Dim obj As UIElement = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(control, obj) Then
                _positionMap(control) = New Point(x, y)
                GraphicsWindow.BeginInvoke(
                    Sub()
                        obj.BeginAnimation(Canvas.LeftProperty, Nothing)
                        obj.BeginAnimation(Canvas.TopProperty, Nothing)
                        Canvas.SetLeft(obj, x)
                        Canvas.SetTop(obj, y)
                    End Sub)
            End If
        End Sub

        ''' <summary>
        ''' Sets the size of the control.
        ''' </summary>
        ''' <param name="control">
        ''' The name of the control to be resized.
        ''' </param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        Public Shared Sub SetSize(control As Primitive, width As Primitive, height As Primitive)
            Dim value As UIElement = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(control, value) Then
                Return
            End If
            Dim element As FrameworkElement = TryCast(value, FrameworkElement)
            If element IsNot Nothing Then
                GraphicsWindow.BeginInvoke(
                    Sub()
                        element.Width = width
                        element.Height = height
                    End Sub)
            End If
        End Sub

        ''' <summary>
        ''' Hides an already added control.
        ''' </summary>
        ''' <param name="controlName">
        ''' The name of the control.
        ''' </param>
        Public Shared Sub HideControl(controlName As Primitive)
            Dim obj As UIElement = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(controlName, obj) Then
                GraphicsWindow.Invoke(Sub() obj.Visibility = Visibility.Collapsed)
            End If
        End Sub

        ''' <summary>
        ''' Shows a previously hidden control.
        ''' </summary>
        ''' <param name="controlName">
        ''' The name of the control.
        ''' </param>
        Public Shared Sub ShowControl(controlName As Primitive)
            Dim obj As UIElement = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(controlName, obj) Then
                GraphicsWindow.Invoke(
                    Sub() obj.Visibility = Visibility.Visible
                )
            End If
        End Sub

        Private Shared Sub OnButtonClicked(sender As Object, e As RoutedEventArgs)
            Dim button1 = TryCast(sender, Button)
            Dim name = GW_NAME + "." + button1.Name
            If GraphicsWindow._objectsMap.ContainsKey(name) Then
                _lastClickedButton = name
                RaiseEvent ButtonClicked()
            End If
        End Sub

        Private Shared Sub OnTextChanged(sender As Object, e As EventArgs)
            Dim textBox1 = TryCast(sender, TextBox)
            Dim name = GW_NAME + "." + textBox1.Name
            If GraphicsWindow._objectsMap.ContainsKey(name) Then
                _lastTypedTextBox = name
                RaiseEvent TextTyped()
            End If
        End Sub
    End Class
End Namespace
