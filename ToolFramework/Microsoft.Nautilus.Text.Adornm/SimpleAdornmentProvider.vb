Imports System.Collections.Generic
Imports Microsoft.Nautilus.Text.AdornmentSystem

Namespace Microsoft.Nautilus.Text.Adornments
    Friend NotInheritable Class SimpleAdornmentProvider
        Implements ISimpleProvider, IAdornmentProvider

        Private _adornments As New List(Of IAdornment)
        Private _textBuffer As ITextBuffer

        Public Event AdornmentsChanged As EventHandler(Of AdornmentsChangedEventArgs) Implements IAdornmentProvider.AdornmentsChanged

        Public Function GetAdornments(span As SnapshotSpan) As IList(Of IAdornment) Implements ISimpleProvider.GetAdornments
            If span.Snapshot.TextBuffer IsNot _textBuffer Then
                Throw New ArgumentException("span")
            End If

            Dim list1 As New List(Of IAdornment)
            For Each adornment As IAdornment In _adornments
                If adornment.Span.GetSpan(span.Snapshot).OverlapsWith(span) Then
                    list1.Add(adornment)
                End If
            Next

            Return list1.AsReadOnly()
        End Function

        Public Function AddAdornment(Of T As IAdornment)(adornment As T) As T Implements ISimpleProvider.AddAdornment
            If adornment.Span.TextBuffer IsNot _textBuffer Then
                Throw New ArgumentException("adornment")
            End If

            _adornments.Add(adornment)
            RaiseAdornmentsChangedEvent(adornment.Span)
            Return adornment
        End Function

        Public Function RemoveAdornment(adornment As IAdornment) As Boolean Implements ISimpleProvider.RemoveAdornment
            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If

            Dim span As ITextSpan = adornment.Span
            If _adornments.Remove(adornment) Then
                RaiseAdornmentsChangedEvent(span)
                Return True
            End If
            Return False
        End Function

        Friend Sub New(textBuffer As ITextBuffer)
            If textBuffer Is Nothing Then
                Throw New ArgumentNullException("textBuffer")
            End If
            _textBuffer = textBuffer
        End Sub

        Private Sub RaiseAdornmentsChangedEvent(span As ITextSpan)
            RaiseEvent AdornmentsChanged(Me, New AdornmentsChangedEventArgs(span))
        End Sub
    End Class
End Namespace
