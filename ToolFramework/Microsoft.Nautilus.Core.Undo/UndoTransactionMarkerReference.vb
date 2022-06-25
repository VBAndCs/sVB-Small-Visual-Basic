
Namespace Microsoft.Nautilus.Core.Undo
    Friend Class UndoTransactionMarkerReferenceValuePair
        Private transactionRef As WeakReference

        Public ReadOnly Property Transaction As UndoTransactionImpl
            Get
                Dim result As UndoTransactionImpl = Nothing

                Try
                    result = TryCast(transactionRef.Target, UndoTransactionImpl)
                    Return result

                Catch __unusedInvalidOperationException1__ As InvalidOperationException
                    Return result
                End Try
            End Get
        End Property

        Public ReadOnly Property Value As Object

        Public Sub New(transaction As UndoTransactionImpl, value As Object)
            transactionRef = New WeakReference(transaction)
            _Value = value
        End Sub
    End Class
End Namespace
