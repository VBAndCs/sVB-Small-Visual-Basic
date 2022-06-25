
Namespace Microsoft.Nautilus.Core.Undo
    Friend Class AutoEnclose
        Implements IDisposable

        Private [end] As AutoEncloseDelegate

        Public Sub New([end] As AutoEncloseDelegate)
            Me.end = [end]
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Call [end]()
        End Sub
    End Class
End Namespace
