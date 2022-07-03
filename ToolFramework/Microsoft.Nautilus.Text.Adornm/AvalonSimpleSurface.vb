Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Adornments
    Friend Class AvalonSimpleSurface
        Implements IAdornmentSurface

        Private Class MyCanvas
            Inherits Canvas

            Private _surface As AvalonSimpleSurface

            Public Sub New(surface As AvalonSimpleSurface)
                _surface = surface
            End Sub

            Protected Overrides Sub OnRender(dc As DrawingContext)
                MyBase.OnRender(dc)
                _surface.OnRender(dc)
            End Sub
        End Class

        Private _canvas As MyCanvas
        Private _textView As IAvalonTextView
        Private _avalonVisualFactories As IEnumerable(Of ImportInfo(Of AvalonVisualFactory, IAvalonVisualFactoryMetadata))
        Private _adornmentsAndVisuals As New Dictionary(Of SimpleAdornment, AvalonVisual)
        Friend _factories As New Dictionary(Of String, ImportInfo(Of AvalonVisualFactory, IAvalonVisualFactoryMetadata))

        Public ReadOnly Property CanNegotiateSurfaceSpace As Boolean = False Implements IAdornmentSurface.CanNegotiateSurfaceSpace

        Public ReadOnly Property SurfacePosition As SurfacePosition = SurfacePosition.BelowText Implements IAdornmentSurface.SurfacePosition

        Public ReadOnly Property SurfacePanel As Panel Implements IAdornmentSurface.SurfacePanel
            Get
                Return _canvas
            End Get
        End Property

        Public Sub AddAdornment(adornment As IAdornment) Implements IAdornmentSurface.AddAdornment
            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If

            If TypeOf adornment IsNot SimpleAdornment Then
                Throw New NotSupportedException("Only adornments of type SimpleAdornment are supported by this surface.")
            End If

            Dim simpleAdorn = CType(adornment, SimpleAdornment)
            Dim visual As AvalonVisual = Nothing

            If Not _adornmentsAndVisuals.TryGetValue(simpleAdorn, visual) Then
                visual = CreateVisual(simpleAdorn)
                If visual IsNot Nothing Then
                    _adornmentsAndVisuals.Add(simpleAdorn, visual)
                    visual.OnExposed(_textView, simpleAdorn)
                    _canvas.InvalidateVisual()
                End If
            End If
        End Sub

        Public Sub RemoveAdornment(adornment As IAdornment) Implements IAdornmentSurface.RemoveAdornment
            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If

            If TypeOf adornment IsNot SimpleAdornment Then
                Throw New NotSupportedException("Only adornments of type SimpleAdornment are supported by this surface.")
            End If

            Dim simpleAdornment1 = CType(adornment, SimpleAdornment)
            Dim _value As AvalonVisual = Nothing

            If _adornmentsAndVisuals.TryGetValue(simpleAdornment1, _value) Then
                _adornmentsAndVisuals.Remove(simpleAdornment1)
                _value.OnHidden(_textView, simpleAdornment1)
                _canvas.InvalidateVisual()
            End If
        End Sub

        Public Function GetSpaceNegotiations(textLine As ITextLine) As IList(Of SpaceNegotiation) Implements IAdornmentSurface.GetSpaceNegotiations
            If textLine Is Nothing Then
                Throw New ArgumentNullException("textLine")
            End If

            Dim lineSpan = textLine.LineSpan
            Dim results As New List(Of SpaceNegotiation)

            For Each adorn In _adornmentsAndVisuals
                Dim key = adorn.Key
                If lineSpan.Contains(key.Span.GetSpan(_textView.TextSnapshot)) Then
                    Dim negotiations = adorn.Value.GetSpaceNegotiations(_textView, key, textLine)
                    If negotiations IsNot Nothing Then
                        results.AddRange(negotiations)
                    End If
                End If
            Next

            If results.Count <> 0 Then Return results

            Return Nothing
        End Function

        Public Function GetAdornmentsGeometry() As Geometry Implements IAdornmentSurface.GetAdornmentsGeometry
            Return Nothing
        End Function

        Friend Sub TextView_LayoutChanged(sender As Object, e As TextViewLayoutChangedEventArgs)
            If Not e.ChangeSpan.HasValue Then Return

            Dim span = e.ChangeSpan.Value
            For Each adornmentsAndVisual In _adornmentsAndVisuals
                Dim adorn = adornmentsAndVisual.Key
                If adorn.Span.GetSpan(_textView.TextSnapshot).OverlapsWith(span) Then
                    adornmentsAndVisual.Value.OnChanged(_textView, adorn)
                End If
            Next
            _canvas.InvalidateVisual()
        End Sub

        Friend Sub New(textView As IAvalonTextView, avalonVisualFactories As IEnumerable(Of ImportInfo(Of AvalonVisualFactory, IAvalonVisualFactoryMetadata)))
            _textView = textView
            _avalonVisualFactories = avalonVisualFactories
            _canvas = New MyCanvas(Me)
            AddHandler _textView.LayoutChanged, AddressOf TextView_LayoutChanged
        End Sub

        Private Function CreateVisual(adornment As SimpleAdornment) As AvalonVisual
            Dim adornmentType = adornment.Type
            Dim _value As ComponentModel.Composition.ImportInfo(Of AvalonVisualFactory, IAvalonVisualFactoryMetadata) = Nothing

            If Not _factories.TryGetValue(adornmentType, _value) Then
                _value = Nothing
                For Each avalonVisualFactory1 As ImportInfo(Of AvalonVisualFactory, IAvalonVisualFactoryMetadata) In _avalonVisualFactories
                    Dim visualFactoryType1 = avalonVisualFactory1.Metadata.VisualFactoryType
                    If visualFactoryType1 = adornmentType Then
                        _value = avalonVisualFactory1
                        Exit For
                    End If
                Next
                _factories.Add(adornmentType, _value)
            End If
            Return _value?.GetBoundValue().CreateAvalonVisual(adornment)
        End Function

        Friend Sub OnRender(dc As DrawingContext)
            For Each item In _adornmentsAndVisuals
                item.Value.OnRender(_textView, item.Key, dc)
            Next
        End Sub
    End Class
End Namespace
