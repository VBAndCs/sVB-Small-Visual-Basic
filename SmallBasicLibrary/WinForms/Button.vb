Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication

Namespace WinForms
    <SmallBasicType>
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
                        GetText = Label.GetTextBlock(buttonName).Text
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

    End Class
End Namespace