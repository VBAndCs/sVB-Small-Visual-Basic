Imports System.Windows
Imports System.Windows.Media

Module Helper

    Public Function GetParent(Of ParentType)(child As DependencyObject) As ParentType
        If child Is Nothing Then Return Nothing
        Dim p = child
        Dim t = GetType(ParentType)
        Do
            p = VisualTreeHelper.GetParent(p)
            If p Is Nothing Then Return Nothing
            If p.GetType Is t Then Return CType(CObj(p), ParentType)
        Loop
    End Function
End Module
