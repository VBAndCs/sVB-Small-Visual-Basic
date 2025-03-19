Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media

Namespace WinForms
    Public Module extensions
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
        Public Function GetChild(Of T)(parent As UIElement, Optional recurse As Boolean = False) As T
            Library.Internal.SmallBasicApplication.Invoke(
            Sub()
                For Each c In parent.GetChildren(recurse)
                    If TypeOf c Is T Then
                        GetChild = CObj(c)
                        Return
                    End If
                Next
            End Sub)
        End Function

        <System.Runtime.CompilerServices.Extension()>
        Public Function GetChild(parent As FrameworkElement, name As String) As FrameworkElement
            For Each c In parent.GetChildren()
                Dim fw = TryCast(c, FrameworkElement)
                If fw?.Name = name Then Return fw
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