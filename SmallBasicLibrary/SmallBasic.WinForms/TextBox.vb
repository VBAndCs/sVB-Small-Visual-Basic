Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class TextBox

        Private Shared Function GetTextBox(formName As String, textBoxName As String) As Wpf.TextBox
            Dim c = Control.GetControl(formName, textBoxName)
            Dim t = TryCast(c, Wpf.TextBox)
            If t Is Nothing Then
                MsgBox($"{textBoxName} is not a name of a TextBox.")
            End If
            Return t
        End Function

        <ExProperty>
        Public Shared Function GetText(formName As Primitive, textBoxName As Primitive) As Primitive
            App.Invoke(Sub() GetText = GetTextBox(formName, textBoxName).Text)
        End Function

        <ExProperty>
        Public Shared Sub SetText(formName As Primitive, textBoxName As Primitive, value As Primitive)
            App.Invoke(Sub() GetTextBox(formName, textBoxName).Text = value)
        End Sub

        <ExProperty>
        Public Shared Function GetMuliLine(formName As Primitive, textBoxName As Primitive) As Primitive
            App.Invoke(Sub() GetMuliLine = GetTextBox(formName, textBoxName).AcceptsReturn)
        End Function

        <ExProperty>
        Public Shared Sub SetMuliLine(formName As Primitive, textBoxName As Primitive, value As Primitive)
            App.Invoke(Sub() GetTextBox(formName, textBoxName).AcceptsReturn = CBool(value))
        End Sub


        <ExMethod>
        Public Shared Sub [Select](formName As Primitive, controlName As Primitive, startPos As Primitive, length As Primitive)
            App.Invoke(Sub()
                           Dim txt = GetTextBox(formName, controlName)
                           Dim st = Math.Max(0, startPos - 1)
                           Dim en = 0
                           If length > 0 Then
                               en = Math.Min(length, txt.Text.Length - st)
                           End If
                           txt.Select(st, en)
                       End Sub)
        End Sub

        <ExMethod>
        Public Shared Sub SelectAll(formName As Primitive, controlName As Primitive)
            App.Invoke(Sub() GetTextBox(formName, controlName).SelectAll())
        End Sub

        Public Shared Custom Event OnTextChanged As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Dim VisualElement = GetTextBox([Event].SenderForm, [Event].SenderControl)
                AddHandler VisualElement.TextChanged, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(Sender, e, handler)
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        Public Shared Custom Event OnTextInput As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Dim VisualElement = GetTextBox([Event].SenderForm, [Event].SenderControl)

                ' Wpf doesn't raise TextInput when space is pressed!
                ' We will fix this by raising this special case from the KeyDown event
                AddHandler VisualElement.PreviewKeyDown,
                    Sub(Sender As Wpf.Control, e As System.Windows.Input.KeyEventArgs)
                        If e.Key = Keys.Space Then
                            Keyboard._lastTextInput = " "
                            [Event].EventsHandler(Sender, e, handler)
                        End If
                    End Sub

                AddHandler VisualElement.PreviewTextInput, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(Sender, e, handler)
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace