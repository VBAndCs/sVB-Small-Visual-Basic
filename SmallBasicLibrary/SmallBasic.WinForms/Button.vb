Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallBasic.Library.Internal.SmallBasicApplication

Namespace WinForms
    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class Button

        Private Shared Function GetButton(formName As String, buttonName As String) As Wpf.Button
            Dim c = Control.GetControl(formName, buttonName)
            Dim t = TryCast(c, Wpf.Button)
            If t Is Nothing Then
                MsgBox($"{buttonName} is not a name of a Button.")
            End If
            Return t
        End Function

        ''' <summary>
        ''' Gets or sets the test that is displayed on the button
        ''' </summary>
        <ExProperty>
        Public Shared Function GetText(formName As Primitive, buttonName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetText = GetButton(formName, buttonName).Content.ToString()
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, buttonName, "Text", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetText(formName As Primitive, buttonName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetButton(formName, buttonName).Content = CStr(value)
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, buttonName, "Text", value, ex.Message)
                    End Try
                End Sub)
        End Sub
    End Class
End Namespace