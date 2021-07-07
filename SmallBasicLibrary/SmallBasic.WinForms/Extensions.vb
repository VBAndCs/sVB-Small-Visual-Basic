Imports System.Windows
Imports System.Windows.Media

Namespace WinForms
    Module Extentions
        <System.Runtime.CompilerServices.Extension()>
        Public Iterator Function GetChildren(ByVal parent As UIElement, ByVal Optional recurse As Boolean = True) As IEnumerable(Of UIElement)
            If parent IsNot Nothing Then
                Dim count As Integer = VisualTreeHelper.GetChildrenCount(parent)

                For i As Integer = 0 To count - 1
                    Dim child = TryCast(VisualTreeHelper.GetChild(parent, i), UIElement)

                    If child IsNot Nothing Then
                        Yield child

                        If recurse Then
                            For Each grandChild In child.GetChildren(True)
                                Yield grandChild
                            Next
                        End If
                    End If
                Next
            End If
        End Function

    End Module
End Namespace