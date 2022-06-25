
Namespace Microsoft.Nautilus.Text
    Public Class TextChangedEventArgs
        Inherits EventArgs

        Public ReadOnly Property Before As ITextSnapshot

        Public ReadOnly Property After As ITextSnapshot

        Public ReadOnly Property Changes As INormalizedTextChangeCollection

        Public ReadOnly Property SourceToken As Object

        Public Sub New(
                    beforeSnapshot As ITextSnapshot,
                    afterSnapshot As ITextSnapshot,
                    changes As INormalizedTextChangeCollection,
                    sourceToken As Object
         )

            _Before = beforeSnapshot
            _After = afterSnapshot
            _Changes = changes
            _SourceToken = sourceToken
        End Sub

    End Class
End Namespace
