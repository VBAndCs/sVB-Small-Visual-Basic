Imports System.Threading

Namespace Microsoft.Nautilus.Core.Task
    Public Class ControllableAsyncTask
        Inherits AsyncTask

        Public Sub AsyncTaskBeginInvoke(taskWorker1 As TaskWorker, taskWorkerState As Object, taskWorkerSyncContext As SynchronizationContext, finishedCallback As AsyncTaskFinishedCallback, finishedCallbackSyncContext As SynchronizationContext)
            MyBase.BeginInvoke(taskWorker1, taskWorkerState, taskWorkerSyncContext, finishedCallback, finishedCallbackSyncContext)
        End Sub
    End Class

    Public Class ControllableAsyncTask(Of T)
        Inherits AsyncTask(Of T)

        Public Sub AsyncTaskBeginInvoke(
                    taskWorker As TaskWorker(Of T),
                    taskWorkerState As Object,
                    taskWorkerSyncContext As SynchronizationContext,
                    finishedCallback As AsyncTaskFinishedCallback,
                    finishedCallbackSyncContext As SynchronizationContext
        )

            MyBase.BeginInvoke(taskWorker, taskWorkerState, taskWorkerSyncContext, finishedCallback, finishedCallbackSyncContext)
        End Sub
    End Class
End Namespace
