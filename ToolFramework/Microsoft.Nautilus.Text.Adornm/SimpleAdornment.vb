Imports Microsoft.Nautilus.Text.AdornmentSystem

Namespace Microsoft.Nautilus.Text.Adornments
    Public Class SimpleAdornment
        Implements IAdornment

        Public ReadOnly Property Span As ITextSpan Implements IAdornment.Span

        Public ReadOnly Property Type As String

        Public ReadOnly Property Subtype As Object

        Public Sub New(span As ITextSpan, type As String, subtype As Object)
            If span Is Nothing Then
                Throw New ArgumentNullException("span")
            End If

            If type Is Nothing Then
                Throw New ArgumentNullException("type")
            End If

            _Span = span
            _Type = type
            _Subtype = subtype
        End Sub
    End Class
End Namespace
