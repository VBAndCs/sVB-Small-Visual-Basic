Namespace Microsoft.Nautilus.Core.Task
    Public Delegate Function TaskWorker(state As Object, task As CancelableTask) As TaskWorkerResult
    Public Delegate Function TaskWorker(Of T)(state As Object, task As CancelableTask) As TaskWorkerResult(Of T)
End Namespace
