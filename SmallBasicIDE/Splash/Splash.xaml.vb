Imports System

Public Class Splash



    Public Shared Sub ShowSplash()
        Dim splash As New Splash()
        splash.Show()
        Dim close As New RunAction(Sub() splash.Close())
        close.After(3000)
    End Sub
End Class
