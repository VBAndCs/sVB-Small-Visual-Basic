Imports SmallBasicLibrary.Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls

<SmallBasicType>
Public NotInheritable Class Button

    Private Shared Function GetButton(formName As String, buttonName As String) As Wpf.Button
        Dim c = Control.GetControl(formName, buttonName)
        Dim t = TryCast(c, Wpf.Button)
        If t Is Nothing Then
            Throw New ArgumentException($"{buttonName} is not a name of a Button.")
        End If
        Return t
    End Function

    <ExProperty>
    Public Shared Function GetText(formName As Primitive, buttonName As Primitive) As Primitive
        Forms.Dispatcher.Invoke(Sub() GetText = GetButton(formName, buttonName).Content.ToString())
    End Function

    <ExProperty>
    Public Shared Sub SetText(formName As Primitive, buttonName As Primitive, value As Primitive)
        Forms.Dispatcher.Invoke(Sub() GetButton(formName, buttonName).Content = CStr(value))
    End Sub
End Class