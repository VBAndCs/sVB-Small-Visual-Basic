Imports System.Threading

Namespace Microsoft.Nautilus.Core.Task
    Public MustInherit Class AsyncTask
        Inherits CancelableTask

        Protected ReadOnly Property [SyncLock] As Object

        Protected Property BeginInvokeRequested As Boolean

        Protected Property EndInvokeRequested As Boolean

        Protected Property AsyncTaskFinishedEvent As ManualResetEvent

        Protected Property TaskWorkerFinishedEvent As ManualResetEvent

        Protected Property ExceptionEncountered As Exception

        Public Overridable ReadOnly Property AsyncWaitHandle As WaitHandle

        Protected Sub New()
            _BeginInvokeRequested = False
            _EndInvokeRequested = False
            _AsyncTaskFinishedEvent = New ManualResetEvent(initialState:=False)
            _TaskWorkerFinishedEvent = New ManualResetEvent(initialState:=False)
        End Sub

        Friend Overridable Sub BeginInvoke(taskWorker As TaskWorker, taskWorkerState As Object, taskWorkerSyncContext As SynchronizationContext, finishedCallback As AsyncTaskFinishedCallback, finishedCallbackSyncContext As SynchronizationContext)
            If taskWorker Is Nothing Then
                Throw New ArgumentNullException("taskWorker")
            End If

            If taskWorkerSyncContext Is Nothing Then
                Throw New ArgumentNullException("taskWorkerSyncContext")
            End If

            SyncLock _SyncLock
                If _BeginInvokeRequested Then
                    Throw New InvalidOperationException("BeginInvoke has already been called.")
                End If

                _BeginInvokeRequested = True
                taskWorkerSyncContext.Post(
                    Sub()
                        State = TaskState.Running
                        Dim taskWorkerResult1 As TaskWorkerResult = taskWorker(taskWorkerState, Me)
                        State = taskWorkerResult1.State
                        _ExceptionEncountered = taskWorkerResult1.ExceptionEncountered
                        _TaskWorkerFinishedEvent.Set()
                        If finishedCallback IsNot Nothing Then
                            If finishedCallbackSyncContext IsNot Nothing Then
                                finishedCallbackSyncContext.Send(Sub()
                                                                     finishedCallback(State)
                                                                 End Sub, Nothing)
                            Else
                                finishedCallback(State)
                            End If
                        End If
                        _AsyncTaskFinishedEvent.Set()
                    End Sub, Nothing)
            End SyncLock
        End Sub

        Public Overridable Sub EndInvoke()
            SyncLock _SyncLock
                If _EndInvokeRequested Then
                    Throw New InvalidOperationException("EndInvoke had already been called.")
                End If

                _EndInvokeRequested = True
                _TaskWorkerFinishedEvent.WaitOne()
                If _ExceptionEncountered IsNot Nothing Then
                    Throw _ExceptionEncountered
                End If
            End SyncLock
        End Sub
    End Class

    Public MustInherit Class AsyncTask(Of T)
        Inherits AsyncTask
        Private result As T

        Friend Overridable Overloads Sub BeginInvoke(taskWorker As TaskWorker(Of T), taskWorkerState As Object, taskWorkerSyncContext As SynchronizationContext, finishedCallback As AsyncTaskFinishedCallback, finishedCallbackSyncContext As SynchronizationContext)
            If taskWorker Is Nothing Then
                Throw New ArgumentNullException("taskWorker")
            End If

            If taskWorkerSyncContext Is Nothing Then
                Throw New ArgumentNullException("taskWorkerSyncContext")
            End If

            SyncLock MyBase.SyncLock
                If MyBase.BeginInvokeRequested Then
                    Throw New InvalidOperationException("BeginInvoke has already been called.")
                End If

                MyBase.BeginInvokeRequested = True
                taskWorkerSyncContext.Post(
                    Sub()
                        State = TaskState.Running
                        Dim taskWorkerResult1 As TaskWorkerResult(Of T) = taskWorker(taskWorkerState, Me)
                        State = taskWorkerResult1.State
                        MyBase.ExceptionEncountered = taskWorkerResult1.ExceptionEncountered

                        If MyBase.ExceptionEncountered Is Nothing Then
                            result = taskWorkerResult1.Result
                        End If

                        MyBase.TaskWorkerFinishedEvent.Set()
                        If finishedCallback IsNot Nothing Then
                            If finishedCallbackSyncContext IsNot Nothing Then
                                finishedCallbackSyncContext.Send(
                                   Sub()
                                       finishedCallback(State)
                                   End Sub,
                                   Nothing)
                            Else
                                finishedCallback(State)
                            End If
                        End If

                        MyBase.AsyncTaskFinishedEvent.Set()
                    End Sub, Nothing)
            End SyncLock
        End Sub

        Public Overridable Shadows Function EndInvoke() As T
            SyncLock MyBase.SyncLock
                If MyBase.EndInvokeRequested Then
                    Throw New InvalidOperationException("EndInvoke had already been called.")
                End If

                MyBase.EndInvokeRequested = True
                MyBase.TaskWorkerFinishedEvent.WaitOne()
                If MyBase.ExceptionEncountered IsNot Nothing Then
                    Throw MyBase.ExceptionEncountered
                End If

                Return result
            End SyncLock
        End Function
    End Class

End Namespace
