Namespace Microsoft.Nautilus.Text
    Friend Class TextVersionImpl
        Inherits TextVersion

        Private normalizedChanges As INormalizedTextChangeCollection

        Public Overrides ReadOnly Property [Next] As TextVersion

        Public Overrides ReadOnly Property Changes As INormalizedTextChangeCollection
            Get
                Return normalizedChanges
            End Get
        End Property

        Public Sub New(textBuffer As ITextBuffer, versionNumber As Integer)
            MyBase.New(textBuffer, versionNumber)
        End Sub

        Public Function CreateNext(changes As INormalizedTextChangeCollection) As TextVersionImpl
            normalizedChanges = changes
            Dim nxt = New TextVersionImpl(MyBase.TextBuffer, MyBase.VersionNumber + 1)
            _Next = nxt
            Return nxt
        End Function

    End Class
End Namespace
