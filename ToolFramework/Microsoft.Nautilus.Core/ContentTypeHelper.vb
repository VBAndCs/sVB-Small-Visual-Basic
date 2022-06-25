Imports Microsoft.Contracts

Namespace Microsoft.Nautilus.Core
    Public NotInheritable Class ContentTypeHelper

        Public Shared Function GetParentContentType(contentType As String) As String
            Contract.RequiresNotNull(contentType, "contentType")
            Dim num As Integer = contentType.LastIndexOf("."c)
            If num <> -1 Then
                Return contentType.Substring(0, num)
            End If
            Return ""
        End Function

        Public Shared Function IsRootContentType(contentType As String) As Boolean
            Contract.RequiresNotNull(contentType, "contentType")
            Return contentType.IndexOf("."c) = -1
        End Function

        Public Shared Function IsOfType(descendant As String, ancestor As String) As Boolean
            Contract.RequiresNotNull(descendant, "descendant")
            Contract.RequiresNotNull(ancestor, "ancestor")
            If Not IsSame(descendant, ancestor) Then
                Return descendant.ToLowerInvariant().Trim().StartsWith(ancestor.ToLowerInvariant().Trim() & ".")
            End If
            Return True
        End Function

        Public Shared Function IsSame(contentType1 As String, contentType2 As String) As Boolean
            Contract.RequiresNotNull(contentType1, "contentType1")
            Contract.RequiresNotNull(contentType2, "contentType2")
            Return contentType1.ToLowerInvariant().Trim() = contentType2.ToLowerInvariant().Trim()
        End Function

    End Class
End Namespace
