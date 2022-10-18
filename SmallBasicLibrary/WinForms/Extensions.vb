Imports System.Windows
Imports System.Windows.Media

Namespace WinForms
    Public Module Extentions
        <System.Runtime.CompilerServices.Extension()>
        Public Iterator Function GetChildren(parent As UIElement, Optional recurse As Boolean = True) As IEnumerable(Of UIElement)
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

        <System.Runtime.CompilerServices.Extension()>
        Public Function GetChild(Of T)(parent As UIElement) As UIElement
            For Each c In parent.GetChildren()
                If TypeOf c Is T Then Return c
            Next
            Return Nothing
        End Function
    End Module


    Public Class ExPropertyAttribute
        Inherits Attribute

    End Class

    Public Class ExMethodAttribute
        Inherits Attribute

    End Class

    Public Class ReturnValueTypeAttribute
        Inherits Attribute

        Public ReturnTypeValue As VariableType

        Public Sub New(type As VariableType)
            ReturnTypeValue = type
        End Sub
    End Class

End Namespace