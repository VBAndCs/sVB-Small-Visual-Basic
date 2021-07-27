Imports System.Windows
Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls

Namespace WinForms

    <SmallBasicType>
    Public NotInheritable Class [Event]

        ''' <summary>
        ''' Gets or sets the name of the form that raised the event
        ''' </summary>
        ''' <remarks>
        ''' If you are adding an event handler manually, You must set Event.SenderForm and Event.SenderControl before adding an event handler, or just call `Event.HandleEventsOf(form, control)` to do this.
        ''' But it is easier to add event handlers from the upper lists in the code editor, so the designer hides all these details in the .sb.gen file.
        ''' If you use one sub to handle an event of more than one control, read SenderForm and SenderControl to get info about which contol is firing the event now.
        ''' </remarks>
        Public Shared Property SenderForm As Primitive

        ''' <summary>
        ''' Gets or sets the name of the control that raised the event
        ''' </summary>
        ''' <remarks>
        ''' If you are adding an event handler manually, You must set Event.SenderForm and Event.SenderControl before adding an event handler, or just call `Event.HandleEventsOf(form, control)` to do this.
        ''' But it is easier to add event handlers from the upper lists in the code editor, so the designer hides all these details in the .sb.gen file.
        ''' If you use one sub to handle an event of more than one control, read SenderForm and SenderControl to get info about which contol is firing the event now.
        ''' </remarks>
        Public Shared Property SenderControl As Primitive

        ''' <summary>
        ''' If you set this property to true inside an event handler sub, the event will be considered handled and will not be processed bey windows anymore.
        ''' For exanple, if you want to cancel writrting a key to the textbox, use `Event.Handled = True` inside OnKeyDown handler 
        ''' </summary>
        ''' <returns></returns>
        Public Shared Property Handled As Primitive

        Public Shared ReadOnly Property LastKey As Primitive
            Get
                Return Keyboard.LastKey
            End Get
        End Property

        Public Shared ReadOnly Property LastTextInput As Primitive
            Get
                Return Keyboard.LastTextInput
            End Get
        End Property

        Public Shared ReadOnly Property LastKeyName As Primitive
            Get
                Return Keyboard.LastKeyName
            End Get
        End Property

        ''' <summary>
        ''' Set the Evemt.SenderForm and Event.SenderControl. 
        ''' You must call this method before adding event handlers for contrl events.
        ''' </summary>
        ''' <param name="FormName"></param>
        ''' <param name="ControlName"></param>
        Public Shared Sub HandleEventsOf(FormName As Primitive, ControlName As Primitive)
            _SenderForm = FormName
            _SenderControl = ControlName
        End Sub

        Shared Sub EventsHandler(Sender As Wpf.Control, e As RoutedEventArgs, userEventHandler As SmallBasicCallback)
            If e.Source IsNot Sender Then Return

            _SenderControl = Sender.Name
            If TypeOf Sender Is Window Then
                _SenderForm = Sender.Name
            Else
                _SenderForm = CType(Control.GetParent(Sender, GetType(Window)), Window).Name
            End If

            Call userEventHandler()

            ' the handler may set the Handled property. We will use it and reset it.
            e.Handled = _Handled
            _Handled = False
        End Sub

    End Class

End Namespace