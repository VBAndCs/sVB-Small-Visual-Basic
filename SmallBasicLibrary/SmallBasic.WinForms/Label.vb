Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallBasic.Library.Internal.SmallBasicApplication

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class Label

        Private Shared Function GetLabel(formName As String, labelName As String) As Wpf.Label
            Dim c = Control.GetControl(formName, labelName)
            Dim t = TryCast(c, Wpf.Label)
            If t Is Nothing Then
                Throw New Exception($"{labelName} is not a name of a Label.")
            End If
            Return t
        End Function

        <ExProperty>
        Public Shared Function GetText(formName As Primitive, labelName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetText = GetLabel(formName, labelName).Content.ToString()
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, labelName, "Text", ex.Message)
                    End Try
                End Sub)
        End Function


        <ExProperty>
        Public Shared Sub SetText(formName As Primitive, labelName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetLabel(formName, labelName).Content = CStr(value)
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, labelName, "Text", value, ex.Message)
                    End Try
                End Sub)
        End Sub
    End Class
End Namespace