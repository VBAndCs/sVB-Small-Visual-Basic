Imports System.Diagnostics

Namespace Microsoft.Contracts
    Public NotInheritable Class Contract

        Public Shared Sub RequiresNotNull(argument As Object, argumentName As String)
            If argument Is Nothing Then
                Throw New ArgumentNullException(argumentName)
            End If
        End Sub

        <Conditional("DEBUG")>
        Public Shared Sub Assume(expression As Boolean)
        End Sub

        Public Shared Sub Requires(e As Exception)
            If e IsNot Nothing Then
                Throw e
            End If
        End Sub
    End Class
End Namespace
