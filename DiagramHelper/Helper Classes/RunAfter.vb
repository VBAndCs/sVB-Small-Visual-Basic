Imports System.Windows.Threading

Public Class RunAction
    Dim WithEvents Timer As New DispatcherTimer()

    Dim RunAction As Action

    Public Sub New()

    End Sub

    Public Sub New(afterMilliseconds As Integer, action As Action)
        Me.Timer.Interval = TimeSpan.FromMilliseconds(afterMilliseconds)
        Me.RunAction = action
    End Sub

    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer.Tick
        Timer.Stop()
        RunAction.Invoke()
    End Sub

    Public Sub Start()
        Me.Timer.Start()
    End Sub

    Public ReadOnly Property Started As Boolean
        Get
            Return Me.Timer.IsEnabled
        End Get
    End Property

    Public Shared Sub After(afterMilliseconds As Integer, action As Action)
        Dim ra As New RunAction(afterMilliseconds, action)
        ra.Start()
    End Sub


End Class
