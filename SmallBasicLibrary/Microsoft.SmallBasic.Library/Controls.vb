Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls

Namespace Microsoft.SmallBasic.Library
    ''' <summary>
    ''' The Controls object allows you to add, move and interact with controls.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Controls
        Private Shared _positionMap As New Dictionary(Of String, Point)
        Private Shared _lastClickedButton As Primitive
        Private Shared _lastTypedTextBox As Primitive
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
        Public Shared Custom Event ButtonClicked As SmallBasicCallback
            AddHandler(Value As SmallBasicCallback)
                Dim h = TryCast(Events("ButtonClicked"), SmallBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler("ButtonClicked", h)
                Events.AddHandler("ButtonClicked", Value)
            End AddHandler

            RemoveHandler(Value As SmallBasicCallback)
                Events.RemoveHandler("ButtonClicked", Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events("ButtonClicked"), SmallBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Raises an event when text is typed into any TextBox control.
        ''' </summary>
        Public Shared Custom Event TextTyped As SmallBasicCallback
            AddHandler(Value As SmallBasicCallback)
                Dim h = TryCast(Events("TextTyped"), SmallBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler("TextTyped", h)
                Events.AddHandler("TextTyped", Value)
            End AddHandler

            RemoveHandler(Value As SmallBasicCallback)
                Events.RemoveHandler("TextTyped", Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events("TextTyped"), SmallBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent
        End Event
        ''' <summary>
        ''' Adds a button to the graphics window at the specified position.
        ''' </summary>
        ''' <param name="caption">
        ''' The caption to display in the button.
        ''' </param>
        ''' <param name="left">
        ''' The x co-ordinate of the button.
        ''' </param>
        ''' <param name="top">
        ''' The y co-ordinate of the button.
        ''' </param>
        ''' <returns>
        ''' The button that was just added to the Graphics Window.
        ''' </returns>
        Public Shared Function AddButton(caption As Primitive, left As Primitive, top As Primitive) As Primitive
            Dim name As String = Shapes.GenerateNewName("Button")
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
        ''' <param name="buttonName">
        ''' The Button whose caption is requested.
        ''' </param>
        ''' <returns>
        ''' The current caption of the button.
        ''' </returns>
        Public Shared Function GetButtonCaption(buttonName As Primitive) As Primitive
            Dim value As Object = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(buttonName, value) Then
                Return ""
            End If
            Dim button1 As Button = TryCast(value, Button)
            If button1 Is Nothing Then
                Return ""
            End If

            Return GraphicsWindow.InvokeWithReturn(Function() CType(button1.Content.ToString(), Primitive))
        End Function

        ''' <summary>
        ''' Sets the caption of the specified button.
        ''' </summary>
        ''' <param name="buttonName">
        ''' The Button whose caption needs to be set.
        ''' </param>
        ''' <param name="caption">
        ''' The new caption for the button.
        ''' </param>
        Public Shared Sub SetButtonCaption(buttonName As Primitive, caption As Primitive)
            Dim value As Object = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(buttonName, value) Then
                Return
            End If
            Dim button1 As Button = TryCast(value, Button)
            If button1 IsNot Nothing Then
                GraphicsWindow.BeginInvoke(Sub()
                                               button1.Content = caption
                                           End Sub)
            End If
        End Sub
        ''' <summary>
        ''' Adds a text input box to the graphics window at the specified position.
        ''' </summary>
        ''' <param name="left">
        ''' The x co-ordinate of the text box.
        ''' </param>
        ''' <param name="top">
        ''' The y co-ordinate of the text box.
        ''' </param>
        ''' <returns>
        ''' The text box that was just added to the Graphics Window.
        ''' </returns>
        Public Shared Function AddTextBox(left As Primitive, top As Primitive) As Primitive
            Dim name As String = Shapes.GenerateNewName("TextBox")
            GraphicsWindow.Invoke(Sub()
                                      Dim textBox1 As New TextBox With
                                      {
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
        ''' Adds a multi-line text input box to the graphics window at the specified position.
        ''' </summary>
        ''' <param name="left">
        ''' The x co-ordinate of the text box.
        ''' </param>
        ''' <param name="top">
        ''' The y co-ordinate of the text box.
        ''' </param>
        ''' <returns>
        ''' The text box that was just added to the Graphics Window.
        ''' </returns>
        Public Shared Function AddMultiLineTextBox(left As Primitive, top As Primitive) As Primitive
            Dim name As String = Shapes.GenerateNewName("TextBox")
            GraphicsWindow.Invoke(Sub()
                                      Dim textBox1 As New TextBox With
                                      {
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
            Dim value As Object = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(textBoxName, value) Then
                Return ""
            End If
            Dim textBox1 As TextBox = TryCast(value, TextBox)
            If textBox1 Is Nothing Then
                Return ""
            End If

            Return CType(GraphicsWindow.InvokeWithReturn(Function() CType(textBox1.Text, Primitive)), Primitive)
        End Function
        ''' <summary>
        ''' Sets the text of the specified TextBox.
        ''' </summary>
        ''' <param name="textBoxName">
        ''' The TextBox whose text needs to be set.
        ''' </param>
        ''' <param name="text">
        ''' The new text for the TextBox.
        ''' </param>
        Public Shared Sub SetTextBoxText(textBoxName As Primitive, text1 As Primitive)
            Dim value As Object = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(textBoxName, value) Then
                Return
            End If
            Dim textBox1 As TextBox = TryCast(value, TextBox)
            If textBox1 IsNot Nothing Then
                GraphicsWindow.BeginInvoke(Sub()
                                               textBox1.Text = text1
                                           End Sub)
            End If
        End Sub
        ''' <summary>
        ''' Removes a control from the Graphics Window.
        ''' </summary>
        ''' <param name="controlName">
        ''' The name of the control that needs to be removed.
        ''' </param>
        Public Shared Sub Remove(controlName As Primitive)
            Dim value As Object = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(controlName, value) Then
                Dim button1 As Button = TryCast(value, Button)
                If button1 IsNot Nothing Then
                    RemoveHandler button1.Click, AddressOf OnButtonClicked
                End If
                GraphicsWindow.RemoveShape(controlName)
            End If
        End Sub
        ''' <summary>
        ''' Moves the control with the specified name to a new position.
        ''' </summary>
        ''' <param name="control">
        ''' The name of the control to move.
        ''' </param>
        ''' <param name="x">
        ''' The x co-ordinate of the new position.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the new position.
        ''' </param>
        Public Shared Sub Move(control As Primitive, x As Primitive, y As Primitive)
            Dim obj As Object = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(control, obj) Then
                _positionMap(control) = New Point(x, y)
                GraphicsWindow.BeginInvoke(Sub()
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
        ''' <param name="width">
        ''' The width of the control.
        ''' </param>
        ''' <param name="height">
        ''' The height of the control.
        ''' </param>
        Public Shared Sub SetSize(control As Primitive, width1 As Primitive, height1 As Primitive)
            Dim value As Object = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(control, value) Then
                Return
            End If
            Dim element As FrameworkElement = TryCast(value, FrameworkElement)
            If element IsNot Nothing Then
                GraphicsWindow.BeginInvoke(Sub()
                                               element.Width = width1
                                               element.Height = height1
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
            Dim obj As Object = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(controlName, obj) Then
                GraphicsWindow.Invoke(Sub()
                                          obj.Visibility = Visibility.Collapsed
                                      End Sub)
            End If
        End Sub
        ''' <summary>
        ''' Shows a previously hidden control.
        ''' </summary>
        ''' <param name="controlName">
        ''' The name of the control.
        ''' </param>
        Public Shared Sub ShowControl(controlName As Primitive)
            Dim obj As Object = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(controlName, obj) Then
                GraphicsWindow.Invoke(Sub()
                                          obj.Visibility = Visibility.Visible
                                      End Sub)
            End If
        End Sub
        Private Shared Sub OnButtonClicked(sender As Object, e As RoutedEventArgs)
            Dim button1 As Button = TryCast(sender, Button)
            Dim name As String = button1.Name
            If GraphicsWindow._objectsMap.ContainsKey(name) Then
                _lastClickedButton = name
                RaiseEvent ButtonClicked()
            End If
        End Sub
        Private Shared Sub OnTextChanged(sender As Object, e As EventArgs)
            Dim textBox1 As TextBox = TryCast(sender, TextBox)
            Dim name As String = textBox1.Name
            If GraphicsWindow._objectsMap.ContainsKey(name) Then
                _lastTypedTextBox = name
                RaiseEvent TextTyped()
            End If
        End Sub
    End Class
End Namespace
