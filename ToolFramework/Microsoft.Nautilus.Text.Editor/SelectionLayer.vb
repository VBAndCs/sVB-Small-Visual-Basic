Imports System.Collections
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class SelectionLayer
        Inherits Canvas
        Implements ITextSelection, IEnumerable(Of ITextSpan), IEnumerable

        Private _selectionBrush As Brush
        Private _selectionPen As Pen
        Private _avalonTextView As AvalonTextView
        Private _textSpanList As List(Of ITextSpan)
        Private _geometryLookup As Dictionary(Of ITextSpan, Geometry)

        Public Property ActiveSpan As ITextSpan Implements ITextSelection.ActiveSpan
            Get
                If IsEmpty Then
                    Return _avalonTextView.TextSnapshot.CreateTextSpan(0, 0, SpanTrackingMode.EdgeExclusive)
                End If

                Return _textSpanList(_textSpanList.Count - 1)
            End Get

            Set(value As ITextSpan)
                If value Is Nothing Then
                    Throw New ArgumentNullException("value")
                End If

                If value.TextBuffer IsNot _avalonTextView.TextBuffer Then
                    Throw New ArgumentException("value")
                End If

                EnsureTextSpanSpansAcrossTextElement(value)
                If IsEmpty Then
                    _textSpanList.Add(value)
                Else
                    _textSpanList(_textSpanList.Count - 1) = value
                End If

                RaiseEvent SelectionChanged(Me, New EventArgs)
            End Set
        End Property

        Public Property ActiveSnapshotSpan As SnapshotSpan Implements ITextSelection.ActiveSnapshotSpan
            Get
                Return ActiveSpan.GetSpan(_avalonTextView.TextSnapshot)
            End Get

            Set(value As SnapshotSpan)
                If value.Snapshot.TextBuffer IsNot _avalonTextView.TextBuffer Then
                    Throw New ArgumentException("TextBuffer in the given SnapshotSpan does not match with the TextBuffer in the AvalonTextView.")
                End If

                ActiveSpan = value.Snapshot.CreateTextSpan(value, SpanTrackingMode.EdgeExclusive)
            End Set
        End Property

        Public ReadOnly Property Count As Integer Implements ITextSelection.Count
            Get
                Return _textSpanList.Count
            End Get
        End Property

        Public Property IsActiveSpanReversed As Boolean Implements ITextSelection.IsActiveSpanReversed

        Public ReadOnly Property IsEmpty As Boolean Implements ITextSelection.IsEmpty
            Get
                For Each textSpan As ITextSpan In _textSpanList
                    If Not textSpan.GetSpan(_avalonTextView.TextSnapshot).IsEmpty Then
                        Return False
                    End If
                Next

                Return True
            End Get
        End Property

        Default Public ReadOnly Property Item(index As Integer) As ITextSpan Implements ITextSelection.Item
            Get
                Return _textSpanList(index)
            End Get
        End Property

        Public Event SelectionChanged As EventHandler Implements ITextSelection.SelectionChanged

        Private Sub EnsureTextSpanSpansAcrossTextElement(ByRef textSpan As ITextSpan)
            Dim span1 As Span = textSpan.GetSpan(_avalonTextView.TextSnapshot)
            If Not span1.IsEmpty Then
                Dim textElementSpan As Span = _avalonTextView.GetTextElementSpan(span1.Start)
                Dim textElementSpan2 As Span = _avalonTextView.GetTextElementSpan(span1.End - 1)
                If span1.Start <> textElementSpan.Start OrElse span1.End <> textElementSpan2.End Then
                    textSpan = _avalonTextView.TextSnapshot.CreateTextSpan(textElementSpan.Start, textElementSpan2.End - textElementSpan.Start, SpanTrackingMode.EdgeExclusive)
                End If
            End If
        End Sub

        Public Sub New(textView As AvalonTextView)
            _avalonTextView = textView
            MyBase.SnapsToDevicePixels = True
            _textSpanList = New List(Of ITextSpan)
            _geometryLookup = New Dictionary(Of ITextSpan, Geometry)

            Dim highlightColor1 As Color = SystemColors.HighlightColor
            Color.FromArgb(96, highlightColor1.R, highlightColor1.G, highlightColor1.B)

            Dim color1 As Color = Color.FromArgb(48, highlightColor1.R, highlightColor1.G, highlightColor1.B)
            _selectionBrush = New SolidColorBrush(color1)
            _selectionPen = New Pen(SystemColors.HighlightBrush, 1.0)
            _selectionPen.LineJoin = PenLineJoin.Round

            If _selectionBrush.CanFreeze Then
                _selectionBrush.Freeze()
            End If
            If _selectionPen.CanFreeze Then
                _selectionPen.Freeze()
            End If

            AddHandler _avalonTextView.LayoutChanged, AddressOf TextView_LayoutChanged
            AddHandler SelectionChanged, AddressOf SelectionLayer_SelectionChanged
        End Sub

        Public Sub [Select](span As ITextSpan) Implements ITextSelection.Select
            If span Is Nothing Then
                Throw New ArgumentException("span")
            End If

            EnsureTextSpanSpansAcrossTextElement(span)
            _textSpanList.Add(span)
            RaiseEvent SelectionChanged(Me, New EventArgs)
        End Sub

        Public Sub Clear() Implements ITextSelection.Clear
            _textSpanList.Clear()
            RaiseEvent SelectionChanged(Me, New EventArgs)
        End Sub

        Public Function GetEnumerator() As IEnumerator(Of ITextSpan) Implements Collections.Generic.IEnumerable(Of ITextSpan).GetEnumerator
            Return _textSpanList.GetEnumerator()
        End Function

        Private Function GetEnumerator1() As IEnumerator Implements Collections.IEnumerable.GetEnumerator
            Return CType(_textSpanList, IEnumerable).GetEnumerator()
        End Function

        Protected Overrides Sub OnRender(drawingContext1 As DrawingContext)
            MyBase.OnRender(drawingContext1)
            For Each textSpan As ITextSpan In _textSpanList
                Dim geometry1 As Geometry = _geometryLookup(textSpan)
                If geometry1 IsNot Nothing Then
                    drawingContext1.DrawGeometry(_selectionBrush, _selectionPen, geometry1)
                End If
            Next
        End Sub

        Private Sub UpdateGeometries()
            _geometryLookup.Clear()
            For Each textSpan As ITextSpan In _textSpanList
                Dim markerGeometry As Geometry = _avalonTextView.SpanGeometry.GetMarkerGeometry(textSpan.GetSpan(_avalonTextView.TextSnapshot))
                _geometryLookup(textSpan) = markerGeometry
            Next
        End Sub

        Private Sub TextView_LayoutChanged(sender As Object, e As EventArgs)
            UpdateGeometries()
            InvalidateVisual()
        End Sub

        Private Sub SelectionLayer_SelectionChanged(sender As Object, e As EventArgs)
            UpdateGeometries()
            InvalidateVisual()
        End Sub

    End Class
End Namespace
