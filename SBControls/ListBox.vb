Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls

<SmallBasicType>
Public NotInheritable Class ListBox

    Private Shared Function GetListBox(formName As String, ListBoxName As String) As Wpf.ListBox
        Dim c = Control.GetControl(formName, ListBoxName)
        Dim t = TryCast(c, Wpf.ListBox)
        If t Is Nothing Then
            Throw New ArgumentException($"{ListBoxName} is not a name of a ListBox.")
        End If
        Return t
    End Function

    Public Shared Function GetCount(formName As Primitive, ListBoxName As Primitive) As Primitive
        Forms.Dispatcher.Invoke(Sub() GetCount = GetListBox(formName, ListBoxName).Items.Count)
    End Function


    Public Shared Sub Add(formName As Primitive, ListBoxName As Primitive, item As Primitive)
        Forms.Dispatcher.Invoke(Sub() GetListBox(formName, ListBoxName).Items.Add(item))
    End Sub

End Class