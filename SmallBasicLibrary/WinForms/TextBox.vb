Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    ''' <summary>
    ''' Represents the TextBox control, which allows the user to input text.
    ''' Use the Text property to read the text the user wrote.
    ''' Use the OnTextInput event to intercept the text before it is written to the textbox.
    ''' Use the OnTextChanged to take action after the text of the textbox changes.
    ''' You can use the form designer to add a textbox to the form by dragging it from the toolbox.
    ''' It is also possible to use the Form.AddTextBox method to create a new textbox and add it to the form at runtime.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class TextBox

        Friend Shared Function GetTextBox(textBoxName As String) As Wpf.TextBox
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
        Public Shared Function GetMultiLine(textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetMultiLine = GetTextBox(textBoxName).AcceptsReturn
                    Catch ex As Exception
                        Control.ReportError(textBoxName, "MultiLine", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetMultiLine(textBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetTextBox(textBoxName).AcceptsReturn = CBool(value)
                    Catch ex As Exception
                        Control.ReportPropertyError(textBoxName, "MultiLine", value, ex)
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

                    Catch ex As Exception
                        Control.ReportSubError(textBoxName, "AppendLine", ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Gets or sets whether or not to wrap words to the next line if they exceed the textbox width.
        ''' When set this property to False, the horizontal scroll bar will appear when any line exceeds the textbox width.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetWordWrap(textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetWordWrap = (GetTextBox(textBoxName).TextWrapping <> TextWrapping.NoWrap)
                    Catch ex As Exception
                        Control.ReportError(textBoxName, "WordWrap", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetWordWrap(textBoxName As Primitive, Value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           GetTextBox(textBoxName).TextWrapping = If(CBool(Value), TextWrapping.Wrap, TextWrapping.NoWrap)
                       Catch ex As Exception
                           Control.ReportPropertyError(textBoxName, "WordWrap", Value, ex)
                       End Try
                   End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets whether or not to draw a line under the text.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetUnderlined(textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetUnderlined = (GetTextBox(textBoxName).TextDecorations Is TextDecorations.Underline)
                    Catch ex As Exception
                        Control.ReportError(textBoxName, "Underlined", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetUnderlined(textBoxName As Primitive, Value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           GetTextBox(textBoxName).TextDecorations = If(CBool(Value), TextDecorations.Underline, Nothing)
                       Catch ex As Exception
                           Control.ReportPropertyError(textBoxName, "Underlined", Value, ex)
                       End Try
                   End Sub)
        End Sub

        ''' <summary>
        ''' Returns True if you can redo the last action on the TextBox.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetCanRedo(textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetCanRedo = GetTextBox(textBoxName).CanRedo
                    Catch ex As Exception
                        Control.ReportError(textBoxName, "CanRedo", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Returns True if you can undo the last action on the TextBox.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetCanUndo(textBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetCanUndo = GetTextBox(textBoxName).CanUndo
                    Catch ex As Exception
                        Control.ReportError(textBoxName, "CanUndo", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Call this method to Redo the last action on the TextBox
        ''' </summary>
        <ExMethod>
        Public Shared Sub Redo(textBoxName As Primitive)
            App.Invoke(
                 Sub()
                     Try
                         GetTextBox(textBoxName).Redo()
                     Catch ex As Exception
                         Control.ReportSubError(textBoxName, "Redo", ex)
                     End Try
                 End Sub)
        End Sub

        ''' <summary>
        ''' Call this method to undo the last action on the TextBox
        ''' </summary>
        <ExMethod>
        Public Shared Sub Undo(textBoxName As Primitive)
            App.Invoke(
                 Sub()
                     Try
                         GetTextBox(textBoxName).Undo()
                     Catch ex As Exception
                         Control.ReportSubError(textBoxName, "Undo", ex)
                     End Try
                 End Sub)
        End Sub

        ''' <summary>
        ''' Call this method to cut the selected text from the TextBox and add it to the clibboard.
        ''' </summary>
        <ExMethod>
        Public Shared Sub Cut(textBoxName As Primitive)
            App.Invoke(
                 Sub()
                     Try
                         GetTextBox(textBoxName).Cut()
                     Catch ex As Exception
                         Control.ReportSubError(textBoxName, "Cut", ex)
                     End Try
                 End Sub)
        End Sub

        ''' <summary>
        ''' Call this method to copy the selected text from the TextBox to the clibboard.
        ''' </summary>
        <ExMethod>
        Public Shared Sub Copy(textBoxName As Primitive)
            App.Invoke(
                 Sub()
                     Try
                         GetTextBox(textBoxName).Copy()
                     Catch ex As Exception
                         Control.ReportSubError(textBoxName, "Copy", ex)
                     End Try
                 End Sub)
        End Sub

        ''' <summary>
        ''' Call this method to paste the text from the clibboard to the current caret pos in the textbox.
        ''' </summary>
        <ExMethod>
        Public Shared Sub Paste(textBoxName As Primitive)
            App.Invoke(
                 Sub()
                     Try
                         GetTextBox(textBoxName).Paste()
                     Catch ex As Exception
                         Control.ReportSubError(textBoxName, "Paste", ex)
                     End Try
                 End Sub)
        End Sub

        ''' <summary>
        ''' Fired when the text is changed.
        ''' </summary>
        Public Shared Custom Event OnTextChanged As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Try
                    Dim _sender = GetTextBox([Event].SenderControl)
                    AddHandler _sender.TextChanged, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnTextChanged), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired just before text is written to the TextBox. 
        ''' Use Event.LastTextInput to get this text.
        ''' Use Event.Handled = True if you want to cancel writing this text to the TexBox. 
        ''' </summary>
        Public Shared Custom Event OnTextInput As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Try
                    Dim _sender = GetTextBox([Event].SenderControl)

                    ' Wpf doesn't raise TextInput when space is pressed!
                    ' We will fix this by raising this special case from the KeyDown event
                    AddHandler _sender.PreviewKeyDown,
                        Sub(Sender As Wpf.Control, e As System.Windows.Input.KeyEventArgs)
                            If e.Key = Keys.Space Then
                                Keyboard._lastTextInput = " "
                                [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                            End If
                        End Sub

                    AddHandler _sender.PreviewTextInput, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnTextInput), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the selected text is changed in TextBox.
        ''' </summary>
        Public Shared Custom Event OnSelection As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Try
                    Dim _sender = GetTextBox([Event].SenderControl)

                    ' Wpf doesn't raise TextInput when space is pressed!
                    ' We will fix this by raising this special case from the KeyDown event
                    AddHandler _sender.PreviewKeyDown,
                        Sub(Sender As Wpf.Control, e As System.Windows.Input.KeyEventArgs)
                            If e.Key = Keys.Space Then
                                Keyboard._lastTextInput = " "
                                [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                            End If
                        End Sub

                    AddHandler _sender.SelectionChanged, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnSelection), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace