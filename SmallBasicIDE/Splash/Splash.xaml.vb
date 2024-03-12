Imports System

Public Class Splash

    Public Shared Sub ShowSplash()
        Dim splash As New Splash()
        splash.Show()
        Dim close As New Windows.Threading.DispatcherTimer(
             New TimeSpan(0, 0, 3),
             Windows.Threading.DispatcherPriority.Background,
             Sub() splash.Close(),
             Windows.Application.Current.Dispatcher
        )
    End Sub
End Class
