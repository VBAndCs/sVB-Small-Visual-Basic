Imports System.Collections.Generic
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Windows.Controls
    Public Class FindMarkerAdornmentSurface
        Inherits Canvas
        Implements IAdornmentSurface

        Private textView As IAvalonTextView
        Private findMarkerList As New List(Of FindMarker)
        Private findMarkerGeometry As New Dictionary(Of FindMarker, Geometry)

        Public ReadOnly Property CanNegotiateSurfaceSpace As Boolean = False Implements IAdornmentSurface.CanNegotiateSurfaceSpace

        Public ReadOnly Property SurfacePosition As SurfacePosition = SurfacePosition.BelowText Implements IAdornmentSurface.SurfacePosition

        Public ReadOnly Property SurfacePanel As Panel Implements IAdornmentSurface.SurfacePanel
            Get
                Return Me
            End Get
        End Property

        Public Sub New(textView As IAvalonTextView)
            Me.textView = textView
            MyBase.SnapsToDevicePixels = False
            AddHandler Me.textView.LayoutChanged, AddressOf TextView_LayoutChanged
        End Sub

        Public Sub AddAdornment(adornment As IAdornment) Implements IAdornmentSurface.AddAdornment
            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If

            If TypeOf adornment IsNot FindMarker Then
                Throw New NotSupportedException("Only adornments of type FindMarker are supported by this surface.")
            End If

            Dim marker = CType(adornment, FindMarker)
            If Not findMarkerList.Contains(marker) Then
                findMarkerList.Add(marker)
                findMarkerGeometry(marker) = Nothing
            End If

            InvalidateVisual()
        End Sub

        Public Sub RemoveAdornment(adornment As IAdornment) Implements IAdornmentSurface.RemoveAdornment
            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If

            If TypeOf adornment IsNot FindMarker Then
                Throw New NotSupportedException("Only adornments of type FindMarker are supported by this surface.")
            End If

            Dim marker = CType(adornment, FindMarker)
            findMarkerList.Remove(marker)
            findMarkerGeometry.Remove(marker)
            InvalidateVisual()
        End Sub

        Public Function GetSpaceNegotiations(textLine As ITextLine) As IList(Of SpaceNegotiation) Implements IAdornmentSurface.GetSpaceNegotiations
            Return Nothing
        End Function

        Public Function GetAdornmentsGeometry() As Geometry Implements IAdornmentSurface.GetAdornmentsGeometry
            Return Nothing
        End Function

        Private Sub TextView_LayoutChanged(sender As Object, e As TextViewLayoutChangedEventArgs)
            If Not e.ChangeSpan.HasValue Then Return

            For Each marker In findMarkerList
                Dim span = marker.Span.GetSpan(textView.TextSnapshot)
                If e.ChangeSpan.Value.OverlapsWith(span) Then
                    findMarkerGeometry(marker) = textView.SpanGeometry.GetMarkerGeometry(span)
                End If
            Next

            InvalidateVisual()
        End Sub

        Protected Overrides Sub OnRender(dc As DrawingContext)
            MyBase.OnRender(dc)
            For Each marker In findMarkerList
                dc.DrawGeometry(marker.Brush, marker.Pen, findMarkerGeometry(marker))
            Next
        End Sub
    End Class
End Namespace
