Imports System.Windows
Imports System.Windows.Input
Imports Microsoft.SmallBasic.Library
Imports App = Microsoft.SmallBasic.Library.Internal.SmallBasicApplication

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class Keyboard

        Shared Sub New()
            App.Invoke(
                Sub() EventManager.RegisterClassHandler(
                        GetType(Window),
                        UIElement.PreviewKeyDownEvent,
                        New RoutedEventHandler(AddressOf KeyDown))
            )

            App.Invoke(
                Sub() EventManager.RegisterClassHandler(
                        GetType(Window),
                        UIElement.PreviewTextInputEvent,
                        New RoutedEventHandler(AddressOf TextInput))
            )
        End Sub

        Private Shared Sub TextInput(sender As Object, e As System.Windows.Input.TextCompositionEventArgs)
            _LastTextInput = e.Text
        End Sub

        Private Shared Sub KeyDown(sender As Object, e As System.Windows.Input.KeyEventArgs)
            _LastKey = e.Key
        End Sub

        Public Shared ReadOnly Property AltPressed As Primitive
            Get
                App.Invoke(
                      Sub()
                          AltPressed = (Input.Keyboard.Modifiers And ModifierKeys.Alt) > 0
                      End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property CtrlPressed As Primitive
            Get
                App.Invoke(
                     Sub()
                         CtrlPressed = (Input.Keyboard.Modifiers And ModifierKeys.Control) > 0
                     End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property ShiftPressed As Primitive
            Get
                App.Invoke(
                       Sub()
                           ShiftPressed = (Input.Keyboard.Modifiers And ModifierKeys.Shift) > 0
                       End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property WinPressed As Primitive
            Get
                App.Invoke(
                    Sub()
                        WinPressed = (Input.Keyboard.Modifiers And ModifierKeys.Windows) > 0
                    End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property CapsLockOn As Primitive
            Get
                App.Invoke(
                     Sub()
                         CapsLockOn = Input.Keyboard.IsKeyToggled(Key.CapsLock)
                     End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property InsertOn As Primitive
            Get
                App.Invoke(
                       Sub()
                           InsertOn = Input.Keyboard.IsKeyToggled(Key.Insert)
                       End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property ScrollOn As Primitive
            Get
                App.Invoke(
                     Sub()
                         ScrollOn = Input.Keyboard.IsKeyToggled(Key.Scroll)
                     End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property NumLockOn As Primitive
            Get
                App.Invoke(
                     Sub()
                         NumLockOn = Input.Keyboard.IsKeyToggled(Key.NumLock)
                     End Sub)
            End Get
        End Property

        Public Shared ReadOnly Property LastKey As Primitive

        Friend Shared _lastTextInput As Primitive
        Public Shared ReadOnly Property LastTextInput As Primitive
            Get
                Return _lastTextInput
            End Get
        End Property

        Public Shared ReadOnly Property LastKeyName As Primitive
            Get
                Return [Enum].GetName(GetType(Input.Key), CInt(LastKey))
            End Get
        End Property
    End Class
End Namespace