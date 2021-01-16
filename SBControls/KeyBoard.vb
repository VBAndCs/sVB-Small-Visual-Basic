Imports Microsoft.SmallBasic.Library

<SmallBasicType>
Public Module Keyboard

    Sub New()
        EventManager.RegisterClassHandler(
            GetType(Window),
            UIElement.PreviewKeyDownEvent,
            New RoutedEventHandler(AddressOf KeyDown)
        )
    End Sub

    Private Sub KeyDown(sender As Object, e As KeyEventArgs)
        _LastKey = e.Key
    End Sub

    Public ReadOnly Property AltPressed As Primitive
        Get
            Return (Input.Keyboard.Modifiers And ModifierKeys.Alt) > 0
        End Get
    End Property

    Public ReadOnly Property CtrlPressed As Primitive
        Get
            Return (Input.Keyboard.Modifiers And ModifierKeys.Control) > 0
        End Get
    End Property

    Public ReadOnly Property ShiftPressed As Primitive
        Get
            Return (Input.Keyboard.Modifiers And ModifierKeys.Shift) > 0
        End Get
    End Property

    Public ReadOnly Property WinPressed As Primitive
        Get
            Return (Input.Keyboard.Modifiers And ModifierKeys.Windows) > 0
        End Get
    End Property

    Public ReadOnly Property CapsLockOn As Primitive
        Get
            Return Input.Keyboard.IsKeyToggled(Input.Key.CapsLock)
        End Get
    End Property

    Public ReadOnly Property InsertOn As Primitive
        Get
            Return Input.Keyboard.IsKeyToggled(Input.Key.Insert)
        End Get
    End Property

    Public ReadOnly Property ScrollOn As Primitive
        Get
            Return Input.Keyboard.IsKeyToggled(Input.Key.Scroll)
        End Get
    End Property

    Public ReadOnly Property NumLockOn As Primitive
        Get
            Return Input.Keyboard.IsKeyToggled(Input.Key.NumLock)
        End Get
    End Property

    Public ReadOnly Property LastKey As Primitive

    Public ReadOnly Property LastKeyName As Primitive
        Get
            Return [Enum].GetName(GetType(Input.Key), CInt(LastKey))
        End Get
    End Property
End Module
