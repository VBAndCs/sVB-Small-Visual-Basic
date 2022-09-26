Imports System

Namespace Microsoft.SmallVisualBasic.Shell
    Public Class RequestCloseEventArgs
        Inherits EventArgs

        Private itemField As Object

        Public ReadOnly Property Item As Object
            Get
                Return itemField
            End Get
        End Property

        Public Sub New(item As Object)
            If item Is Nothing Then
                Throw New ArgumentNullException("item")
            End If

            itemField = item
        End Sub
    End Class
End Namespace
