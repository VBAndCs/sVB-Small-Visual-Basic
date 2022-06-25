
Namespace Microsoft.Nautilus.Core.Task
    Public Class TaskWorkerResult

        Public ReadOnly Property State As TaskState

        Public ReadOnly Property ExceptionEncountered As Exception

        Public Sub New(state As TaskState)
            If state = TaskState.Initializing OrElse state = TaskState.Running Then
                Throw New ArgumentException("The given state is not a valid finished state.")
            End If

            _State = state
        End Sub

        Public Sub New(state As TaskState, exception As Exception)
            Me.New(state)

            If exception Is Nothing Then
                Throw New ArgumentNullException("_ExceptionEncountered")
            End If

            If state = TaskState.Completed Then
                Throw New ArgumentException("The state cannot be Completed and yet result in an _ExceptionEncountered.")
            End If

            Me._ExceptionEncountered = exception
        End Sub
    End Class

    Public Class TaskWorkerResult(Of T)
        Inherits TaskWorkerResult

        Private taskResult As T

        Public ReadOnly Property Result As T
            Get
                If MyBase.ExceptionEncountered IsNot Nothing Then
                    Throw MyBase.ExceptionEncountered
                End If

                Return taskResult
            End Get
        End Property

        Public Sub New(state As TaskState, exception As Exception)
            MyBase.New(state, exception)
        End Sub

        Public Sub New(state As TaskState, result As T)
            MyBase.New(state)

            taskResult = result
        End Sub
    End Class
End Namespace
