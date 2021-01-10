Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls

<SmallBasicType>
Public Module Button

    Private Function GetButton(formName As String, buttonName As String) As Wpf.Button
        Dim c = GetControl(formName, buttonName)
        Dim t = TryCast(c, Wpf.Button)
        If t Is Nothing Then
            Throw New ArgumentException($"{buttonName} is not a name of a Button.")
        End If
        Return t
    End Function

    Public Function GetText(formName As Primitive, buttonName As Primitive) As Primitive
        Dispatcher.Invoke(Sub() GetText = GetButton(formName, buttonName).Content.ToString())
    End Function


    Public Sub SetText(formName As Primitive, buttonName As Primitive, value As Primitive)
        Dispatcher.Invoke(Sub() GetButton(formName, buttonName).Content = CStr(value))
    End Sub
End Module