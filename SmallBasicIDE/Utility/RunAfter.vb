﻿' Mohammad Hamdy Ghanem

Imports System
Imports System.Windows.Threading

Public Class RunAfter
    Dim WithEvents Timer As New DispatcherTimer()

    Dim runAction As Action

    Public Sub New()

    End Sub

    Public Sub New(action As Action)
        RunAction = action
    End Sub

    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer.Tick
        Timer.Stop()
        RunAction.Invoke()
    End Sub

    Public Sub After(afterMilliseconds As Integer, Optional action As Action = Nothing)
        Timer.Stop()
        If action IsNot Nothing Then runAction = action
        Timer.Interval = TimeSpan.FromMilliseconds(afterMilliseconds)
        Timer.Start()
    End Sub
End Class

Public Class RunAfter(Of T)
    Dim WithEvents Timer As New DispatcherTimer()

    Dim runAction As Action(Of T)

    Public Sub New(action As Action(Of T))
        runAction = action
    End Sub

    Dim data As T

    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer.Tick
        Timer.Stop()
        runAction.Invoke(data)
    End Sub

    Public Sub After(afterMilliseconds As Integer, data As T)
        Timer.Stop()
        Timer.Interval = TimeSpan.FromMilliseconds(afterMilliseconds)
        Timer.Start()
    End Sub
End Class