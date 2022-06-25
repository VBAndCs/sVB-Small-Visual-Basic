
Namespace Microsoft.Nautilus.Text.Diagnostics
    Friend Class TraceProvider
        Public Sub [RaiseEvent](etwEvent1 As EtwEvent, eventType1 As EventType)
        End Sub

        Public Sub Enable()
            Console.WriteLine("Trace enabled")
        End Sub

        Public Sub Disable()
            Console.WriteLine("Trace disabled")
        End Sub
    End Class
End Namespace
