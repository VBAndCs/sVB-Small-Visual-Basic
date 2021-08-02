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
                MsgBox($"{labelName} is not a name of a Label.")
            End If
            Return t
        End Function

        <ExProperty>
        Public Shared Function GetText(formName As Primitive, labelName As Primitive) As Primitive
            App.Invoke(Sub() GetText = GetLabel(formName, labelName).Content.ToString())
        End Function


        <ExProperty>
        Public Shared Sub SetText(formName As Primitive, labelName As Primitive, value As Primitive)
            App.Invoke(Sub() GetLabel(formName, labelName).Content = CStr(value))
        End Sub
    End Class
End Namespace