Imports System.Collections.Generic
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Windows.Controls
    Public Class StatementMarkerAdornmentSurface
        Inherits Canvas
        Implements IAdornmentSurface

        Private textView As IAvalonTextView
        Private statementMarkerList As New List(Of StatementMarker)
        Private statementMarkerGeometry As New Dictionary(Of StatementMarker, Geometry)

        Public ReadOnly Property CanNegotiateSurfaceSpace As Boolean = False Implements IAdornmentSurface.CanNegotiateSurfaceSpace

        Public ReadOnly Property SurfacePosition As SurfacePosition = SurfacePosition.BelowText Implements IAdornmentSurface.SurfacePosition

        Public ReadOnly Property SurfacePanel As Panel Implements IAdornmentSurface.SurfacePanel
            Get
                Return Me
            End Get
        End Property

        Public Sub New(textView As IAvalonTextView)
            Me.textView = textView
            MyBase.SnapsToDevicePixels = True
            AddHandler Me.textView.LayoutChanged, AddressOf TextView_LayoutChanged
        End Sub

        Public Sub AddAdornment(adornment As IAdornment) Implements IAdornmentSurface.AddAdornment
            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If

            If TypeOf adornment IsNot StatementMarker Then
                Throw New NotSupportedException("Only adornments of type StatementMarker are supported by this surface.")
            End If

            Dim marker = CType(adornment, StatementMarker)

            If Not statementMarkerList.Contains(marker) Then
                statementMarkerList.Add(marker)
                statementMarkerGeometry(marker) = Nothing
            End If

            InvalidateVisual()
        End Sub

        Public Sub RemoveAdornment(adornment As IAdornment) Implements IAdornmentSurface.RemoveAdornment
            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If


            If TypeOf adornment IsNot StatementMarker Then
                Throw New NotSupportedException("Only adornments of type StatementMarker are supported by this surface.")
            End If

            Dim marker = CType(adornment, StatementMarker)
            statementMarkerList.Remove(marker)
            statementMarkerGeometry.Remove(marker)

            InvalidateVisual()
        End Sub

        Public Function GetSpaceNegotiations(textLine As ITextLine) As IList(Of SpaceNegotiation) Implements IAdornmentSurface.GetSpaceNegotiations
            Return Nothing
        End Function

        Public Function GetAdornmentsGeometry() As Geometry Implements IAdornmentSurface.GetAdornmentsGeometry
            Return Nothing
        End Function

        Private Sub TextView_LayoutChanged(sender As Object, e As TextViewLayoutChangedEventArgs)
            If Not e.ChangeSpan.HasValue Then
                Return
            End If

            For Each statementMarker1 In statementMarkerList
                Dim span1 = statementMarker1.Span.GetSpan(textView.TextSnapshot)
                If e.ChangeSpan.Value.OverlapsWith(span1) Then
                    Dim markerGeometry = textView.SpanGeometry.GetMarkerGeometry(span1)
                    statementMarkerGeometry(statementMarker1) = markerGeometry
                End If
            Next

            InvalidateVisual()
        End Sub

        Protected Overrides Sub OnRender(dc As DrawingContext)
            MyBase.OnRender(dc)
            For Each marker In statementMarkerList
                dc.DrawGeometry(
                    New SolidColorBrush(Color.FromArgb(64, marker.MarkerColor.R, marker.MarkerColor.G, marker.MarkerColor.B)),
                    New Pen(New SolidColorBrush(marker.MarkerColor), 1.0),
                    statementMarkerGeometry(marker)
                )
            Next
        End Sub
    End Class
End Namespace
