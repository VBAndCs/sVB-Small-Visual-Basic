Imports Microsoft.SmallBasic.Library

<SmallBasicType>
Public Module Keyboard

    Sub New()
        Dispatcher.Invoke(
            Sub() EventManager.RegisterClassHandler(
                        GetType(Window),
                        UIElement.PreviewKeyDownEvent,
                        New RoutedEventHandler(AddressOf KeyDown))
        )
    End Sub

    Private Sub KeyDown(sender As Object, e As KeyEventArgs)
        _LastKey = e.Key
    End Sub

    Public ReadOnly Property AltPressed As Primitive
        Get
            Dispatcher.Invoke(
               Sub()
                   AltPressed = (Input.Keyboard.Modifiers And ModifierKeys.Alt) > 0
               End Sub)
        End Get
    End Property

    Public ReadOnly Property CtrlPressed As Primitive
        Get
            Dispatcher.Invoke(
               Sub()
                   CtrlPressed = (Input.Keyboard.Modifiers And ModifierKeys.Control) > 0
               End Sub)
        End Get
    End Property

    Public ReadOnly Property ShiftPressed As Primitive
        Get
            Dispatcher.Invoke(
               Sub()
                   ShiftPressed = (Input.Keyboard.Modifiers And ModifierKeys.Shift) > 0
               End Sub)
        End Get
    End Property

    Public ReadOnly Property WinPressed As Primitive
        Get
            Dispatcher.Invoke(
               Sub()
                   WinPressed = (Input.Keyboard.Modifiers And ModifierKeys.Windows) > 0
               End Sub)
        End Get
    End Property

    Public ReadOnly Property CapsLockOn As Primitive
        Get
            Dispatcher.Invoke(
               Sub()
                   CapsLockOn = Input.Keyboard.IsKeyToggled(Input.Key.CapsLock)
               End Sub)
        End Get
    End Property

    Public ReadOnly Property InsertOn As Primitive
        Get
            Dispatcher.Invoke(
               Sub()
                   InsertOn = Input.Keyboard.IsKeyToggled(Input.Key.Insert)
               End Sub)
        End Get
    End Property

    Public ReadOnly Property ScrollOn As Primitive
        Get
            Dispatcher.Invoke(
               Sub()
                   ScrollOn = Input.Keyboard.IsKeyToggled(Input.Key.Scroll)
               End Sub)
        End Get
    End Property

    Public ReadOnly Property NumLockOn As Primitive
        Get
            Dispatcher.Invoke(
               Sub()
                   NumLockOn = Input.Keyboard.IsKeyToggled(Input.Key.NumLock)
               End Sub)
        End Get
    End Property

    Public ReadOnly Property LastKey As Primitive

    Public ReadOnly Property LastKeyName As Primitive
        Get
            Return [Enum].GetName(GetType(Input.Key), CInt(LastKey))
        End Get
    End Property
End Module
