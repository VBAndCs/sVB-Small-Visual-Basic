Imports System.Windows
Imports System.Windows.Media
Imports Microsoft.VisualBasic

Module Helper
    Public Function CountLines(str As String) As Integer
        If str = "" Then Return 0

        Dim pos As Integer = -1
        Dim count As Integer = 0
        Do
            pos = str.IndexOf(vbLf, pos + 1)
            If pos = -1 Then Exit Do
            count += 1
        Loop
        Return count
    End Function


    Public Function GetParent(Of ParentType)(child As DependencyObject) As ParentType
        If child Is Nothing Then Return Nothing
        Try
            Dim p = child
            Dim t = GetType(ParentType)
            Do
                p = VisualTreeHelper.GetParent(p)
                If p Is Nothing Then Return Nothing
                If p.GetType Is t Then Return CType(CObj(p), ParentType)
            Loop
        Catch

        End Try
        Return Nothing
    End Function
End Module
