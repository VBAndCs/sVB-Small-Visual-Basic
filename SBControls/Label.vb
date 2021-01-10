Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls

<SmallBasicType>
Public Module Label

    Private Function GetLabel(formName As String, labelName As String) As Wpf.Label
        Dim c = GetControl(formName, labelName)
        Dim t = TryCast(c, Wpf.Label)
        If t Is Nothing Then
            Throw New ArgumentException($"{labelName} is not a name of a Label.")
        End If
        Return t
    End Function

    Public Function GetText(formName As Primitive, labelName As Primitive) As Primitive
        Dispatcher.Invoke(Sub() GetText = GetLabel(formName, labelName).Content.ToString())
    End Function


    Public Sub SetText(formName As Primitive, labelName As Primitive, value As Primitive)
        Dispatcher.Invoke(Sub() GetLabel(formName, labelName).Content = CStr(value))
    End Sub
End Module