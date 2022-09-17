Imports System.Windows
Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls

Namespace WinForms

    ''' <summary>
    ''' Contains info about thee last fired event
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class [Event]
        Shared Sub ShowErrorMessage(eventName As String, ex As Exception)
            ReportError($"Setting the handler for {[Event].SenderControl}.{eventName} caused an error: {vbCrLf}{ex.Message}", ex)
        End Sub

        ''' <summary>
        ''' Gets or sets the name of the control that raised the event
        ''' </summary>
        ''' <remarks>
        ''' If you are adding an event handler manually, You must set Event.SenderForm and Event.SenderControl before adding an event handler, or just call `Event.HandleEventsOf(form, control)` to do this.
        ''' But it is easier to add event handlers from the upper lists in the code editor, so the designer hides all these details in the .sb.gen file.
        ''' If you use one sub to handle an event of more than one control, read SenderForm and SenderControl to get info about which contol is firing the event now.
        ''' </remarks>
        <ReturnValueType(VariableType.Control)>
        Public Shared Property SenderControl As Primitive

        ''' <summary>
        ''' If you set this property to true inside an event handler sub, the event will be considered handled and will not be processed bey windows anymore.
        ''' For exanple, if you want to cancel writrting a key to the textbox, use `Event.Handled = True` inside OnKeyDown handler 
        ''' </summary>
        ''' <returns></returns>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Property Handled As Primitive

        ''' <summary>
        ''' Returns the last Key pressed on the keyboard. 
        ''' Use The Keys enum members to check they key.
        ''' Examle: If Event.LastKey = Keys.A Then
        ''' </summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property LastKey As Primitive
            Get
                Return Keyboard.LastKey
            End Get
        End Property

        ''' <summary>
        ''' returns the last text that was about to be written to the TextBox
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property LastTextInput As Primitive
            Get
                Return Keyboard.LastTextInput
            End Get
        End Property

        ''' <summary>
        ''' Returns the name of the last key pressed on the keyboard.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property LastKeyName As Primitive
            Get
                Return Keyboard.LastKeyName
            End Get
        End Property

        ''' <summary>
        ''' Set the Evemt.SenderForm and Event.SenderControl. 
        ''' You must call this method before adding event handlers for contrl events.
        ''' </summary>
        ''' <param name="ControlName"></param>
        Public Shared Sub HandleEventsOf(ControlName As Primitive)
            _SenderControl = ControlName
        End Sub

        Shared Sub EventsHandler(Sender As FrameworkElement, e As RoutedEventArgs, userEventHandler As SmallBasicCallback, Optional allowTunneling As Boolean = False)
            Try
                If e.Source IsNot Sender Then
                    If TypeOf Sender IsNot Wpf.Label AndAlso Not allowTunneling Then Return

                ElseIf TypeOf Sender Is Wpf.Label Then
                    ' Shape in a lable. Respond only to the shape
                    If CType(Sender, Wpf.Label).Content IsNot Nothing Then Return
                End If

                _SenderControl = GetControlName(Sender)

                Call userEventHandler()

                ' the handler may set the Handled property. We will use it and reset it.
                e.Handled = _Handled
                _Handled = False
            Catch ex As Exception
                ReportError($"The event handler sub `{userEventHandler.Method.Name}` caused this error: {ex.Message}", ex)
            End Try
        End Sub

        Private Shared Function GetControlName(sender As FrameworkElement) As Primitive
            If TypeOf sender Is Window Then
                Return sender.Name.ToLower
            ElseIf TypeOf sender Is wpf.Canvas Then
                Return CType(sender.Parent, Window).Name.ToLower
            Else
                Dim parent As FrameworkElement = sender.Parent
                Do While parent IsNot Nothing
                    If TypeOf parent Is Window Then
                        Return CType(parent, Window).Name.ToLower + "." + sender.Name.ToLower
                    End If
                    parent = parent.Parent
                Loop
            End If

        End Function

        ''' <summary>
        ''' Gets or sets the mouse cursor's x co-ordinate.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        Public Shared ReadOnly Property MouseX As Primitive
            Get
                Return Mouse.MouseX
            End Get
        End Property

        ''' <summary>
        ''' Gets or sets the mouse cursor's y co-ordinate.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        Public Shared ReadOnly Property MouseY As Primitive
            Get
                Return Mouse.MouseY
            End Get
        End Property

        ''' <summary>
        ''' Gets whether or not the left button is pressed.
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared ReadOnly Property IsLeftButtonDown As Primitive
            Get
                Return Mouse.IsLeftButtonDown
            End Get
        End Property

        ''' <summary>
        ''' Gets whether or not the right button is pressed.
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared ReadOnly Property IsRightButtonDown As Primitive
            Get
                Return Mouse.IsRightButtonDown
            End Get
        End Property

    End Class

End Namespace