Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication

Namespace WinForms
    ''' <summary>
    ''' Represents the Button control, that the user can click to perform the task that you provide in the OnClick event handler.
    ''' You can use the form designer to add a button to the form by dragging it from the toolbox.
    ''' It is also possible to use the Form.AddButton method to create a new button and add it to the form at runtime.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class Button

        Private Shared Function GetButton(buttonName As String) As Wpf.Button
            Dim c = Control.GetControl(buttonName)
            Dim t = TryCast(c, Wpf.Button)
            If t Is Nothing Then
                ReportError($"{buttonName} is not a name of a Button.", Nothing)
            End If
            Return t
        End Function

        ''' <summary>
        ''' Gets or sets the test that is displayed on the button
        ''' </summary>
        <ExProperty>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetText(buttonName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetText = New Primitive(Label.GetTextBlock(buttonName).Text)
                    Catch ex As Exception
                        Control.ReportError(buttonName, "Text", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetText(buttonName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Label.GetTextBlock(buttonName).Text = CStr(value)
                    Catch ex As Exception
                        Control.ReportPropertyError(buttonName, "Text", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets whether or not to draw a line under the text.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetUnderlined(buttonName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetUnderlined = (Label.GetTextBlock(buttonName).TextDecorations Is System.Windows.TextDecorations.Underline)
                    Catch ex As Exception
                        Control.ReportError(buttonName, "Underlined", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetUnderlined(buttonName As Primitive, Value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Label.GetTextBlock(buttonName).TextDecorations = If(CBool(Value), System.Windows.TextDecorations.Underline, Nothing)
                       Catch ex As Exception
                           Control.ReportPropertyError(buttonName, "Underlined", Value, ex)
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
                        GetWordWrap = (Label.GetTextBlock(controlName).TextWrapping = System.Windows.TextWrapping.Wrap)
                    Catch ex As Exception
                        Control.ReportError(controlName, "WordWrap", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetWordWrap(controlName As Primitive, Value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Label.GetTextBlock(controlName).TextWrapping = If(Value, System.Windows.TextWrapping.Wrap, System.Windows.TextWrapping.NoWrap)
                       Catch ex As Exception
                           Control.ReportPropertyError(controlName, "WordWrap", Value, ex)
                       End Try
                   End Sub)
        End Sub


        Public Shared ReadOnly IsFlatProperty As System.Windows.DependencyProperty =
                System.Windows.DependencyProperty.RegisterAttached("IsFlat",
                GetType(Boolean), GetType(Button),
                New System.Windows.PropertyMetadata(False))

        ''' <summary>
        ''' Gets or sets whether or not to show the button with a flat style.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetIsFlat(buttonName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetIsFlat = CBool(Control.GetControl(buttonName).GetValue(IsFlatProperty))
                    Catch ex As Exception
                        Control.ReportError(buttonName, "IsFlat", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetIsFlat(buttonName As Primitive, Value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Dim btn = GetButton(buttonName)
                           Dim flat = CBool(Value)
                           btn.SetValue(IsFlatProperty, flat)
                           btn.Style = If(flat,
                                            btn.FindResource(Wpf.ToolBar.ButtonStyleKey),
                                            Nothing)

                       Catch ex As Exception
                           Control.ReportPropertyError(buttonName, "IsFlat", Value, ex)
                       End Try
                   End Sub)
        End Sub

    End Class
End Namespace