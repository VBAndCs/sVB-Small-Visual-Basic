Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class ToggleButton

        Private Shared Function GetToggleButton(toggleButtonName As String) As Wpf.Primitives.ToggleButton
            Dim c = Control.GetControl(toggleButtonName)
            Dim t = TryCast(c, Wpf.Primitives.ToggleButton)
            If t Is Nothing Then
                ReportError($"{toggleButtonName} is not a name of a ToggleButton.", Nothing)
            End If
            Return t
        End Function

        ''' <summary>
        ''' Gets or sets the test that is displayed on the toggleButton
        ''' </summary>
        <ExProperty>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetText(toggleButtonName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetText = Label.GetTextBlock(toggleButtonName).Text
                    Catch ex As Exception
                        Control.ReportError(toggleButtonName, "Text", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetText(toggleButtonName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Label.GetTextBlock(toggleButtonName).Text = CStr(value)
                    Catch ex As Exception
                        Control.ReportPropertyError(toggleButtonName, "Text", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets whether or not to draw a line under the text.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetUnderlined(toggleButtonName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetUnderlined = (Label.GetTextBlock(toggleButtonName).TextDecorations Is System.Windows.TextDecorations.Underline)
                    Catch ex As Exception
                        Control.ReportError(toggleButtonName, "Underlined", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetUnderlined(toggleButtonName As Primitive, Value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Label.GetTextBlock(toggleButtonName).TextDecorations = If(CBool(Value), System.Windows.TextDecorations.Underline, Nothing)
                       Catch ex As Exception
                           Control.ReportPropertyError(toggleButtonName, "Underlined", Value, ex)
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


        Public Shared ReadOnly IsFlatProperty As System.Windows.DependencyProperty =
                System.Windows.DependencyProperty.RegisterAttached("IsFlat",
                GetType(Boolean), GetType(ToggleButton),
                New System.Windows.PropertyMetadata(False))

        ''' <summary>
        ''' Gets or sets whether or not to show the toggleButton with a flat style.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetIsFlat(toggleButtonName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetIsFlat = CBool(Control.GetControl(toggleButtonName).GetValue(IsFlatProperty))
                    Catch ex As Exception
                        Control.ReportError(toggleButtonName, "IsFlat", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetIsFlat(toggleButtonName As Primitive, Value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Dim btn = GetToggleButton(toggleButtonName)
                           Dim flat = CBool(Value)
                           btn.SetValue(IsFlatProperty, flat)
                           btn.Style = If(flat,
                                            btn.FindResource(Wpf.ToolBar.ToggleButtonStyleKey),
                                            Nothing)

                           btn.Focusable = Not flat

                       Catch ex As Exception
                           Control.ReportPropertyError(toggleButtonName, "IsFlat", Value, ex)
                       End Try
                   End Sub)
        End Sub

        ''' <summary>
        ''' When True, the toggle button is checked.
        ''' When False, the toggle button is unchecked.
        ''' When empty string "",  the toggle button is not checked nor unchecked (indeterminate state)
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetChecked(controlName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim rd = GetToggleButton(controlName)
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
                        Dim rd = GetToggleButton(controlName)
                        rd.IsChecked = CBool(value)
                    Catch ex As Exception
                        Control.ReportPropertyError(controlName, "Checked", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Fired when the checked state is changed.
        ''' </summary>
        Public Shared Custom Event OnCheck As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Try
                    Dim _sender = GetToggleButton([Event].SenderControl)
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