Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class CheckBox

        Private Shared Function GetCheckBox(CheckBoxName As String) As Wpf.CheckBox
            Dim c = Control.GetControl(CheckBoxName)
            Dim t = TryCast(c, Wpf.CheckBox)
            If t Is Nothing Then
                Throw New Exception($"{CheckBoxName} is not a name of a CheckBox.")
            End If
            Return t
        End Function

        ''' <summary>
        ''' Gets or sets the text that is displayed by the CheckBox
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetText(CheckBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetText = GetCheckBox(CheckBoxName).Content
                    Catch ex As Exception
                        Control.ReportError(CheckBoxName, "Text", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetText(CheckBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim t = GetCheckBox(CheckBoxName)
                        t.Content = value
                    Catch ex As Exception
                        Control.RepotPropertyError(CheckBoxName, "Text", value, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' When True, the check box is checked.
        ''' When False, the check box is unchecked.
        ''' When empty string "",  the check box is not checked nor unchecked (indeterminate state)
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetChecked(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim ch = GetCheckBox(controlName)
                        Dim state = ch.IsChecked
                        If state.HasValue Then
                            GetChecked = New Primitive(state.Value)
                        Else
                            GetChecked = ""
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
                        Dim ch = GetCheckBox(controlName)
                        ' Don't use if() expression, because it will return a new primitive not Nothing
                        If value.AsString() = "" Then
                            ch.IsChecked = Nothing
                        Else
                            ch.IsChecked = CBool(value)
                        End If
                    Catch ex As Exception
                        Control.RepotPropertyError(controlName, "Checked", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Controls what happen when the user clicks the checkbox to cange its state:
        ''' * When False(the defaust value): the checkbox state goes from checked to unchecked then back to checked and so on.
        ''' * When True, the the checkbox state goes from checked to dimmed (indeterminate) to unchecked then back to checked and so on.
        ''' Note that you can set the dimmed (indeterminate) state from code by setting the Checked property to an empty string "". This will always work regardless the value of AllowTriState, which has effect on the user interaction only.
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetAllowThreeState(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetAllowThreeState = GetCheckBox(controlName).IsThreeState
                    Catch ex As Exception
                        Control.ReportError(controlName, "AllowThreeState", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetAllowThreeState(controlName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetCheckBox(controlName).IsThreeState = CBool(value)
                    Catch ex As Exception
                        Control.RepotPropertyError(controlName, "Checked", value, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Fired when the text is changed.
        ''' </summary>
        Public Shared Custom Event OnCheck As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Try
                    Dim _sender = GetCheckBox([Event].SenderControl)
                    Dim _handler = Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(Sender, e, handler)
                    AddHandler _sender.Checked, _handler
                    AddHandler _sender.Unchecked, _handler
                    AddHandler _sender.Indeterminate, _handler
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnCheck), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace
