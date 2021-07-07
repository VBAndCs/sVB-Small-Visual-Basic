Imports System.Windows
Imports System.Windows.Input
Imports Microsoft.SmallBasic.Library

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class Keyboard

        Shared Sub New()
            Forms.Dispatcher.Invoke(
            Sub() EventManager.RegisterClassHandler(
                        GetType(Window),
                        UIElement.PreviewKeyDownEvent,
                        New RoutedEventHandler(AddressOf KeyDown))
        )
        End Sub

        Private Shared Sub KeyDown(sender As Object, e As System.Windows.Input.KeyEventArgs)
            _LastKey = e.Key
        End Sub

        Public Shared ReadOnly Property AltPressed As Primitive
            Get
                Forms.Dispatcher.Invoke(
                      Sub()
                          AltPressed = (Input.Keyboard.Modifiers And ModifierKeys.Alt) > 0
                      End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property CtrlPressed As Primitive
            Get
                Forms.Dispatcher.Invoke(
                     Sub()
                         CtrlPressed = (Input.Keyboard.Modifiers And ModifierKeys.Control) > 0
                     End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property ShiftPressed As Primitive
            Get
                Forms.Dispatcher.Invoke(
                       Sub()
                           ShiftPressed = (Input.Keyboard.Modifiers And ModifierKeys.Shift) > 0
                       End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property WinPressed As Primitive
            Get
                Forms.Dispatcher.Invoke(
                    Sub()
                        WinPressed = (Input.Keyboard.Modifiers And ModifierKeys.Windows) > 0
                    End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property CapsLockOn As Primitive
            Get
                Forms.Dispatcher.Invoke(
                     Sub()
                         CapsLockOn = Input.Keyboard.IsKeyToggled(Key.CapsLock)
                     End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property InsertOn As Primitive
            Get
                Forms.Dispatcher.Invoke(
                       Sub()
                           InsertOn = Input.Keyboard.IsKeyToggled(Key.Insert)
                       End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property ScrollOn As Primitive
            Get
                Forms.Dispatcher.Invoke(
                     Sub()
                         ScrollOn = Input.Keyboard.IsKeyToggled(Key.Scroll)
                     End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property NumLockOn As Primitive
            Get
                Forms.Dispatcher.Invoke(
                     Sub()
                         NumLockOn = Input.Keyboard.IsKeyToggled(Key.NumLock)
                     End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property LastKey As Primitive

        Public Shared ReadOnly Property LastKeyName As Primitive
            Get
                Return [Enum].GetName(GetType(Input.Key), CInt(LastKey))
            End Get
        End Property
    End Class
End Namespace