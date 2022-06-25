Namespace Microsoft.Nautilus.Core.Task
    Public MustInherit Class CancelableTask

        Public Overridable Property State As TaskState

        Public Overridable ReadOnly Property CancelRequested As Boolean

        Public Overridable Sub RequestCancel()
            _CancelRequested = True
        End Sub
    End Class
End Namespace
