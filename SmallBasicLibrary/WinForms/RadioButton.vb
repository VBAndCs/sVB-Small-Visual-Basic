Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    ''' Represents a RadioButton control, that the user can check or uncheck.
    ''' You can use the Checked property and OnCheck event to respond to the user choices.
    ''' You can use the form designer to add a radio button to the form by dragging it from the toolbox.
    ''' It is also possible to use the Form.AddRadioButton method to create a new radio button and add it to the form at runtime.
    ''' Radio buttons work togethor as a group, where each group can contain only one checked radio button. 
    ''' By default, all radio buttons you add to the form will be grouped together, but if you want to create more than one goup, you can group some radio buttons together by setting the GroupNmae property in each of them to the same group name.
    ''' You can also use the form designer to group radiobuttons by selecting them, right-clicking one of the selected radio buttons, and clicking the Group command from the context menu.
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class RadioButton

        Private Shared Function GetRadioButton(radioButtonName As String) As Wpf.RadioButton
            Dim c = Control.GetControl(radioButtonName)
            Dim rd = TryCast(c, Wpf.RadioButton)
            If rd Is Nothing Then
                Throw New Exception($"{radioButtonName} is not a name of a RadioButton.")
            End If
            Return rd
        End Function

        ''' <summary>
        ''' Gets or sets the text that is displayed by the RadioButton
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetText(radioButtonName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetText = Label.GetTextBlock(radioButtonName).Text
                    Catch ex As Exception
                        Control.ReportError(radioButtonName, "Text", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetText(radioButtonName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Label.GetTextBlock(radioButtonName).Text = value.AsString()
                    Catch ex As Exception
                        Control.ReportPropertyError(radioButtonName, "Text", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the name of the radio buttons group. Radio buttons that belong to the same group can have only one seleected button at a time, so, if the user checks one button, the previously checked button will be unchecked.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetGroupName(radioButtonName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetGroupName = GetRadioButton(radioButtonName).GroupName
                    Catch ex As Exception
                        Control.ReportError(radioButtonName, "GroupName", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetGroupName(radioButtonName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim rd = GetRadioButton(radioButtonName)
                        rd.GroupName = value
                    Catch ex As Exception
                        Control.ReportPropertyError(radioButtonName, "GroupName", value, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' When True, the radio button is checked.
        ''' When False, the radio button is unchecked.
        ''' When empty string "",  the radio button is not checked nor unchecked (indeterminate state)
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetChecked(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim rd = GetRadioButton(controlName)
                        Dim state = rd.IsChecked
                        If state.HasValue Then
                            GetChecked = New Primitive(state.Value)
                        Else
                            GetChecked = New Primitive(False)
                        End If
                    Catch ex As Exception
                        Control.ReportError(controlName, "Checked", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetChecked(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim rd = GetRadioButton(controlName)
                        rd.IsChecked = CBool(value)
                    Catch ex As Exception
                        Control.ReportPropertyError(controlName, "Checked", value, ex)
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
                        GetUnderlined = (Label.GetTextBlock(controlName).TextDecorations Is System.Windows.TextDecorations.Underline)
                    Catch ex As Exception
                        Control.ReportError(controlName, "Underlined", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetUnderlined(controlName As Primitive, Value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Label.GetTextBlock(controlName).TextDecorations = If(CBool(Value), System.Windows.TextDecorations.Underline, Nothing)
                       Catch ex As Exception
                           Control.ReportPropertyError(controlName, "Underlined", Value, ex)
                       End Try
                   End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets whether or not to the text will be continue on the next line if it exceeds the width of the control.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetWordWrap(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetWordWrap = (Label.GetTextBlock(controlName).TextWrapping = TextWrapping.Wrap)
                    Catch ex As Exception
                        Control.ReportError(controlName, "WordWrap", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetWordWrap(controlName As Primitive, Value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Label.GetTextBlock(controlName).TextWrapping = If(Value, TextWrapping.Wrap, TextWrapping.NoWrap)
                       Catch ex As Exception
                           Control.ReportPropertyError(controlName, "WordWrap", Value, ex)
                       End Try
                   End Sub)
        End Sub

        ''' <summary>
        ''' Fired when the checked state is changed.
        ''' </summary>
        Public Shared Custom Event OnCheck As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Try
                    Dim _sender = GetRadioButton([Event].SenderControl)
                    Dim _handler = Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(Sender, e, handler)
                    AddHandler _sender.Checked, _handler
                    AddHandler _sender.Unchecked, _handler
                    AddHandler _sender.Indeterminate, _handler
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnCheck), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace
