
Namespace Microsoft.Nautilus.Text.Document
    Friend Class VacuousBufferManager
        Implements IBufferManager

        Public ReadOnly Property BaseBuffer As ITextBuffer Implements IBufferManager.BaseBuffer

        Public ReadOnly Property SurfaceBuffer As ITextBuffer Implements IBufferManager.SurfaceBuffer
            Get
                Return _BaseBuffer
            End Get
        End Property

        Public Sub New(baseBuffer As ITextBuffer)
            If baseBuffer Is Nothing Then
                Throw New ArgumentNullException("baseBuffer")
            End If

            _BaseBuffer = baseBuffer
        End Sub
    End Class
End Namespace
