Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class TextBox

        Private Shared Function GetTextBox(textBoxName As String) As Wpf.TextBox
            Dim c = Control.GetControl(textBoxName)
            Dim t = TryCast(c, Wpf.TextBox)
            If t Is Nothing Then
                Throw New Exception($"{textBoxName} is not a name of a TextBox.")
            End If
            Return t
        End Function

        ''' <summary>
        ''' Gets or sets the text that is displayed by the TextBox
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetText(textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetText = GetTextBox(textBoxName).Text
                    Catch ex As Exception
                        Control.ReportError(textBoxName, "Text", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetText(textBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim t = GetTextBox(textBoxName)
                        t.Text = value
                        t.VerticalScrollBarVisibility = Wpf.ScrollBarVisibility.Auto
                        t.HorizontalScrollBarVisibility = Wpf.ScrollBarVisibility.Auto
                    Catch ex As Exception
                        Control.ReportPropertyError(textBoxName, "Text", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the start pos of the selected text.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetSelectionStart(textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetSelectionStart = GetTextBox(textBoxName).SelectionStart + 1
                    Catch ex As Exception
                        Control.ReportError(textBoxName, "SelectionStart", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectionStart(textBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetTextBox(textBoxName).SelectionStart = value - 1
                    Catch ex As Exception
                        Control.ReportPropertyError(textBoxName, "SelectionStart", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the length of the selected text.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetSelectionLength(textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetSelectionLength = GetTextBox(textBoxName).SelectionLength
                    Catch ex As Exception
                        Control.ReportError(textBoxName, "SelectionLength", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectionLength(textBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetTextBox(textBoxName).SelectionLength = value
                    Catch ex As Exception
                        Control.ReportPropertyError(textBoxName, "SelectionLength", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the selected text.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetSelectedText(textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetSelectedText = GetTextBox(textBoxName).SelectedText
                    Catch ex As Exception
                        Control.ReportError(textBoxName, "SelectedText", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedText(textBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim t = GetTextBox(textBoxName)
                        t.SelectedText = value
                        t.VerticalScrollBarVisibility = Wpf.ScrollBarVisibility.Auto
                        t.HorizontalScrollBarVisibility = Wpf.ScrollBarVisibility.Auto
                    Catch ex As Exception
                        Control.ReportPropertyError(textBoxName, "SelectedText", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the current caret pos in the TextBox.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetCaretIndex(textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetCaretIndex = GetTextBox(textBoxName).CaretIndex + 1
                    Catch ex As Exception
                        Control.ReportError(textBoxName, "CaretIndex", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetCaretIndex(textBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetTextBox(textBoxName).CaretIndex = value - 1
                    Catch ex As Exception
                        Control.ReportPropertyError(textBoxName, "CaretIndex", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Set this property to True  to allow the user to write more than one line in the TextBox
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetMuliLine(textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetMuliLine = GetTextBox(textBoxName).AcceptsReturn
                    Catch ex As Exception
                        Control.ReportError(textBoxName, "MuliLine", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetMuliLine(textBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetTextBox(textBoxName).AcceptsReturn = CBool(value)
                    Catch ex As Exception
                        Control.ReportPropertyError(textBoxName, "MuliLine", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Selscts a part ot the text deisplayed in the TextBox.
        ''' </summary>
        ''' <param name="startPos">the pos of the first character you want to select</param>
        ''' <param name="length">the number of characters you want to select</param>
        <ExMethod>
        Public Shared Sub [Select](
                        textBoxName As Primitive,
                        startPos As Primitive,
                        length As Primitive
                   )

            App.Invoke(
                Sub()
                    Try
                        Dim txt = GetTextBox(textBoxName)
                        Dim st = Math.Max(0, startPos - 1)
                        Dim en = 0

                        If length > 0 Then
                            en = Math.Min(length, txt.Text.Length - st)
                        End If

                        txt.Select(st, en)

                    Catch ex As Exception
                        Control.ReportSubError(textBoxName, "Select", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Selects all the text displayed in the TextBox
        ''' </summary>
        <ExMethod>
        Public Shared Sub SelectAll(textBoxName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetTextBox(textBoxName).SelectAll()
                    Catch ex As Exception
                        Control.ReportSubError(textBoxName, "SelectAll", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds the given text at the end of the TextBox
        ''' </summary>
        ''' <param name="text">the text that will be written at the end of the TextBox</param>
        <ExMethod>
        Public Shared Sub Append(textBoxName As Primitive, text As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim t = GetTextBox(textBoxName)
                        t.AppendText(text.AsString())
                        t.VerticalScrollBarVisibility = Wpf.ScrollBarVisibility.Auto
                        t.HorizontalScrollBarVisibility = Wpf.ScrollBarVisibility.Auto

                    Catch ex As Exception
                        Control.ReportSubError(textBoxName, "Append", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds the given text at the end ot the TextBox then adds a new line cahracter, so the next text will be written in a new line
        ''' </summary>
        ''' <param name="lineText">the text that will be written at the end of the TextBox followed by a new line character.</param>
        <ExMethod>
        Public Shared Sub AppendLine(textBoxName As Primitive, lineText As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim t = GetTextBox(textBoxName)
                        t.AppendText(lineText.AsString() & vbCrLf)
                        t.VerticalScrollBarVisibility = Wpf.ScrollBarVisibility.Auto
                        t.HorizontalScrollBarVisibility = Wpf.ScrollBarVisibility.Auto

                    Catch ex As Exception
                        Control.ReportSubError(textBoxName, "AppendLine", ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Gets or sets whether or not to draw a line under the text.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetUnderlined(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetUnderlined = (GetTextBox(controlName).TextDecorations Is TextDecorations.Underline)
                    Catch ex As Exception
                        Control.ReportError(controlName, "Underlined", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetUnderlined(controlName As Primitive, Value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           GetTextBox(controlName).TextDecorations = If(CBool(Value), TextDecorations.Underline, Nothing)
                       Catch ex As Exception
                           Control.ReportPropertyError(controlName, "Underlined", Value, ex)
                       End Try
                   End Sub)
        End Sub


        ''' <summary>
        ''' Fired when the text is changed.
        ''' </summary>
        Public Shared Custom Event OnTextChanged As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Try
                    Dim VisualElement = GetTextBox([Event].SenderControl)
                    AddHandler VisualElement.TextChanged, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnTextChanged), ex)
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
                    Dim VisualElement = GetTextBox([Event].SenderControl)

                    ' Wpf doesn't raise TextInput when space is pressed!
                    ' We will fix this by raising this special case from the KeyDown event
                    AddHandler VisualElement.PreviewKeyDown,
                        Sub(Sender As Wpf.Control, e As System.Windows.Input.KeyEventArgs)
                            If e.Key = Keys.Space Then
                                Keyboard._lastTextInput = " "
                                [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                            End If
                        End Sub

                    AddHandler VisualElement.PreviewTextInput, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnTextInput), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace