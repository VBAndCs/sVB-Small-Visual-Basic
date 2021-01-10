Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls

<SmallBasicType>
Public Module TextBox

    Private Function GetTextBox(formName As String, textBoxName As String) As Wpf.TextBox
        Dim c = GetControl(formName, textBoxName)
        Dim t = TryCast(c, Wpf.TextBox)
        If t Is Nothing Then
            Throw New ArgumentException($"{textBoxName} is not a name of a TextBox.")
        End If
        Return t
    End Function

    Public Function GetText(formName As Primitive, textBoxName As Primitive) As Primitive
        Dispatcher.Invoke(Sub() GetText = GetTextBox(formName, textBoxName).Text)
    End Function


    Public Sub SetText(formName As Primitive, textBoxName As Primitive, value As Primitive)
        Dispatcher.Invoke(Sub() GetTextBox(formName, textBoxName).Text = value)
    End Sub
End Module
