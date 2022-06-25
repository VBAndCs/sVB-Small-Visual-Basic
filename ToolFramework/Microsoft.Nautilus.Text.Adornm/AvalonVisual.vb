Imports System.Collections.Generic
Imports System.Windows.Media
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Adornments
    Public MustInherit Class AvalonVisual

        Public Overridable Sub OnExposed(view As IAvalonTextView, adornment As SimpleAdornment)
            If view Is Nothing Then
                Throw New ArgumentNullException("view")
            End If

            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If
        End Sub

        Public Overridable Sub OnHidden(view As IAvalonTextView, adornment As SimpleAdornment)
            If view Is Nothing Then
                Throw New ArgumentNullException("view")
            End If

            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If
        End Sub

        Public Overridable Sub OnChanged(view As IAvalonTextView, adornment As SimpleAdornment)
            If view Is Nothing Then
                Throw New ArgumentNullException("view")
            End If

            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If
        End Sub

        Public Overridable Function GetSpaceNegotiations(view As IAvalonTextView, adornment As SimpleAdornment, textLine As ITextLine) As IList(Of SpaceNegotiation)
            If view Is Nothing Then
                Throw New ArgumentNullException("view")
            End If

            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If

            If textLine Is Nothing Then
                Throw New ArgumentNullException("textLine")
            End If
            Return Nothing
        End Function

        Public Overridable Sub OnRender(view As IAvalonTextView, adornment As SimpleAdornment, dc As DrawingContext)
            If view Is Nothing Then
                Throw New ArgumentNullException("view")
            End If

            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If

            If dc Is Nothing Then
                Throw New ArgumentNullException("dc")
            End If

        End Sub
    End Class
End Namespace
