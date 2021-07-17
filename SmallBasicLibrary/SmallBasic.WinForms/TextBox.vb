Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallBasic.Library.Internal.SmallBasicApplication

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class TextBox

        Private Shared Function GetTextBox(formName As String, textBoxName As String) As Wpf.TextBox
            Dim c = Control.GetControl(formName, textBoxName)
            Dim t = TryCast(c, Wpf.TextBox)
            If t Is Nothing Then
                Throw New ArgumentException($"{textBoxName} is not a name of a TextBox.")
            End If
            Return t
        End Function

        <ExProperty>
        Public Shared Function GetText(formName As Primitive, textBoxName As Primitive) As Primitive
            App.Invoke(Sub() GetText = GetTextBox(formName, textBoxName).Text)
        End Function

        <ExProperty>
        Public Shared Sub SetText(formName As Primitive, textBoxName As Primitive, value As Primitive)
            App.Invoke(Sub() GetTextBox(formName, textBoxName).Text = value)
        End Sub
    End Class
End Namespace