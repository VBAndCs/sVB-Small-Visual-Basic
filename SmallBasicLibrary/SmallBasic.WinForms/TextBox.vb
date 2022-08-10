Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class TextBox

        Private Shared Function GetTextBox(formName As String, textBoxName As String) As Wpf.TextBox
            Dim c = Control.GetControl(formName, textBoxName)
            Dim t = TryCast(c, Wpf.TextBox)
            If t Is Nothing Then
                Throw New Exception($"{textBoxName} is not a name of a TextBox.")
            End If
            Return t
        End Function

        ''' <summary>
        ''' Gets or sets the text that is displayed on the label
        ''' </summary>
        <ExProperty>
        Public Shared Function GetText(formName As Primitive, textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetText = GetTextBox(formName, textBoxName).Text
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, textBoxName, "Text", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetText(formName As Primitive, textBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetTextBox(formName, textBoxName).Text = value
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, textBoxName, "Text", value, ex.Message)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Set this property to True  to allow the user to write more than one line in the TextBox
        ''' </summary>
        <ExProperty>
        Public Shared Function GetMuliLine(formName As Primitive, textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetMuliLine = GetTextBox(formName, textBoxName).AcceptsReturn
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, textBoxName, "MuliLine", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetMuliLine(formName As Primitive, textBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetTextBox(formName, textBoxName).AcceptsReturn = CBool(value)
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, textBoxName, "MuliLine", value, ex.Message)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Selscts a part ot the text deisplayed in the TextBox.
        ''' </summary>
        ''' <param name="startPos">the pos of the first character you want to select</param>
        ''' <param name="length">the number of characters you want to select</param>
        <ExMethod>
        Public Shared Sub [Select](formName As Primitive, textBoxName As Primitive, startPos As Primitive, length As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim txt = GetTextBox(formName, textBoxName)
                        Dim st = Math.Max(0, startPos - 1)
                        Dim en = 0
                        If length > 0 Then
                            en = Math.Min(length, txt.Text.Length - st)
                        End If
                        txt.Select(st, en)
                    Catch ex As Exception
                        Control.ShowSubError(formName, textBoxName, "Select", ex.Message)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Selects all the text displayed in the TextBox
        ''' </summary>
        <ExMethod>
        Public Shared Sub SelectAll(formName As Primitive, textBoxName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetTextBox(formName, textBoxName).SelectAll()
                    Catch ex As Exception
                        Control.ShowSubError(formName, textBoxName, "SelectAll", ex.Message)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Fired when the text is changed.
        ''' </summary>
        Public Shared Custom Event OnTextChanged As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Try
                    Dim VisualElement = GetTextBox([Event].SenderForm, [Event].SenderControl)
                    AddHandler VisualElement.TextChanged, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnTextChanged), ex.Message)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when before text is written to thee TextBox. 
        ''' Use Event.LastTextInput to get this text.
        ''' Use Event.Handled = True if you want to cancel writing this text to the TexBox. 
        ''' </summary>
        Public Shared Custom Event OnTextInput As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Try
                    Dim VisualElement = GetTextBox([Event].SenderForm, [Event].SenderControl)

                    ' Wpf doesn't raise TextInput when space is pressed!
                    ' We will fix this by raising this special case from the KeyDown event
                    AddHandler VisualElement.PreviewKeyDown,
                        Sub(Sender As Wpf.Control, e As System.Windows.Input.KeyEventArgs)
                            If e.Key = Keys.Space Then
                                Keyboard._LastTextInput = " "
                                [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                            End If
                        End Sub

                    AddHandler VisualElement.PreviewTextInput, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnTextInput), ex.Message)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace