﻿Imports System.Windows
Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls

Namespace WinForms

    ''' <summary>
    ''' Contains info about the last fired event
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class [Event]
        Shared Sub ShowErrorMessage(eventName As String, ex As Exception)
            ReportError($"Setting the handler for {[Event].SenderControl}.{eventName} caused an error: {vbCrLf}{ex.Message}", ex)
        End Sub

        Friend Shared _senderControl As Primitive

        ''' <summary>
        '''Gets the name of the control that raised the last event.
        '''It is useful when you use one sub to handle many events, to get info about the contol that fired the event.
        ''' </summary>
        <ReturnValueType(VariableType.Control)>
        Public Shared ReadOnly Property SenderControl As Primitive
            Get
                Return _senderControl
            End Get
        End Property

        ''' <summary>
        ''' If you set this property to true inside an event handler sub, the event will be considered handled and will not be processed by windows anymore.
        ''' For example, if you want to cancel writrting a key to the textbox, use `Event.Handled = True` inside the OnKeyDown event handler.
        ''' </summary>
        ''' <returns></returns>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Property Handled As Primitive

        ''' <summary>
        ''' Returns the last Key pressed on the keyboard. 
        ''' Use The Keys enum members to check they key.
        ''' Example: If Event.LastKey = Keys.A Then
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
        ''' <param name="ControlName">The cntrl you will handle its events.</param>
        <HideFromIntellisense>
        Public Shared Sub HandleEventsOf(ControlName As Primitive)
            _SenderControl = ControlName
        End Sub

        Shared Sub HandleEvent(
                            sender As FrameworkElement,
                            e As RoutedEventArgs,
                            userEventHandler As SmallVisualBasicCallback,
                            Optional allowTunneling As Boolean = False
                    )

            Try
                If e.Source IsNot sender AndAlso
                    Not (TypeOf e.Source Is Wpf.Canvas AndAlso TryCast(sender, Window)?.AllowsTransparency) AndAlso
                    TypeOf sender IsNot Wpf.Label AndAlso
                    Not allowTunneling Then Return

                _SenderControl = GetControlName(sender)

                Call userEventHandler()

                ' the handler may set the Handled property. We will use it and reset it.
                e.Handled = _Handled
                _Handled = False

            Catch ex As Exception
                ReportError($"The event handler sub `{userEventHandler.Method.Name}` caused this error: {ex.Message}", ex)
            End Try
        End Sub

        Friend Shared Function GetControlName(sender As FrameworkElement) As Primitive
            If TypeOf sender Is Window Then
                Return New Primitive(sender.Name.ToLower())
            ElseIf TypeOf sender Is Wpf.Canvas Then
                Return New Primitive(CType(sender.Parent, Window).Name.ToLower())
            Else
                Dim parent As FrameworkElement = sender.Parent
                Do While parent IsNot Nothing
                    If TypeOf parent Is Window Then
                        Return New Primitive(CType(parent, Window).Name.ToLower() + "." + sender.Name.ToLower())
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
        ''' Get a value that indicates the last mouse wheel movement direction:
        '''  * -1 means down.
        '''  * 1 means up.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        Public Shared ReadOnly Property LastMouseWheelDirection As Primitive
            Get
                Return Mouse.LastMouseWheelDirection
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