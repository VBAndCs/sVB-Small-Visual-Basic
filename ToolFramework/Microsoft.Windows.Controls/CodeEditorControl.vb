Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.ComponentModel.Activation
Imports System.ComponentModel.Composition
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Media
Imports Microsoft.Nautilus.Core
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Classification
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations

Namespace Microsoft.Windows.Controls
    Public Class CodeEditorControl
        Inherits ContentControl
        Implements INotifyPropertyChanged

        Private Const _lineNumberMarginName As String = "Avalon Line Number Margin"
        Public Shared ReadOnly TextBufferProperty As DependencyProperty = DependencyProperty.Register("TextBuffer", GetType(ITextBuffer), GetType(CodeEditorControl), New FrameworkPropertyMetadata(AddressOf TextBufferChanged))
        Public Shared ReadOnly ContentTypeProperty As DependencyProperty = DependencyProperty.Register("ContentType", GetType(String), GetType(CodeEditorControl))
        Public Shared ReadOnly HighlightSearchHitsProperty As DependencyProperty = DependencyProperty.Register("HighlightSearchHits", GetType(Boolean), GetType(CodeEditorControl))

        Private _scaleFactor As Double = 1.0
        Private _textView As AvalonTextView
        Private _textViewHost As AvalonTextViewHost
        Private _isLineNumberMarginVisible As Boolean
        Private _marginTransform As New ScaleTransform(1.0, 1.0)
        Private _lineNumberMargin As Canvas

        Public ReadOnly Property ContainsHighlights As Boolean

        Public ReadOnly Property ContainsWordHighlights As Boolean

        <Import>
        Public Property AdornmentAggregatorCache As IAdornmentAggregatorCache

        <Import>
        Public Property AdornmentSurfaceManagerFactory As IAdornmentSurfaceManagerFactory

        <Import>
        Public Property AvalonTextViewMarginFactories As IEnumerable(Of ImportInfo(Of AvalonTextViewMarginFactory, IAvalonTextViewMarginFactoryMetadata))

        <Import>
        Public Property BufferFactory As ITextBufferFactory

        <Import>
        Public Property ClassificationFormatMap As IClassificationFormatMap

        <Import>
        Public Property ClassificationTypeRegistry As IClassificationTypeRegistry

        <Import>
        Public Property ClassifierAggregatorProvider As IClassifierAggregatorProvider

        Public Property ContentType As String
            Get
                Dim text As String = CStr(GetValue(ContentTypeProperty))
                If text Is Nothing Then
                    text = "text"
                    ContentType = text
                End If

                Return text
            End Get

            Set(value As String)
                SetValue(ContentTypeProperty, value)
            End Set
        End Property

        Public ReadOnly Property EditorOperations As IEditorOperations
            Get
                Return EditorOperationsProvider.GetEditorOperations(TextView)
            End Get
        End Property

        <Import>
        Public Property EditorOperationsProvider As IEditorOperationsProvider

        Public Property HighlightSearchHits As Boolean
            Get
                Return CBool(GetValue(HighlightSearchHitsProperty))
            End Get

            Set(value As Boolean)
                SetValue(HighlightSearchHitsProperty, value)
            End Set
        End Property

        Public Property IsLineNumberMarginVisible As Boolean
            Get
                Return _isLineNumberMarginVisible
            End Get

            Set(value As Boolean)
                _isLineNumberMarginVisible = value
                If _textViewHost IsNot Nothing Then
                    Dim textViewMargin As ITextViewMargin = _textViewHost.GetTextViewMargin("Avalon Line Number Margin")
                    If textViewMargin IsNot Nothing Then
                        textViewMargin.MarginVisible = value
                        _lineNumberMargin = TryCast(textViewMargin, Canvas)
                        If _lineNumberMargin IsNot Nothing Then
                            _lineNumberMargin.LayoutTransform = _marginTransform
                        End If
                    End If
                End If

                Notify("IsLineNumberMarginVisible")
            End Set
        End Property

        <Import>
        Public Property KeyboardFilters As IEnumerable(Of KeyboardFilter)

        <Import>
        Public Property MouseBindingFactories As IEnumerable(Of MouseBindingFactory)

        <Import>
        Public Property Orderer As IOrderer

        Public Property ScaleFactor As Double
            Get
                If _textViewHost IsNot Nothing Then
                    Return _textViewHost.ScaleFactor
                End If

                Return _scaleFactor
            End Get

            Set(value As Double)
                _scaleFactor = value
                If _textViewHost IsNot Nothing Then
                    _textViewHost.ScaleFactor = value
                    _marginTransform.ScaleX = value
                    _marginTransform.ScaleY = value
                End If

                Notify("ScaleFactor")
            End Set
        End Property

        Public Property Text As String
            Get
                Return TextBuffer.CurrentSnapshot.GetText(0, TextBuffer.CurrentSnapshot.Length)
            End Get

            Set(value As String)
                Dim currentSnapshot1 As ITextSnapshot = TextBuffer.CurrentSnapshot
                TextBuffer.Replace(New Span(0, currentSnapshot1.Length), value)
            End Set
        End Property

        Public Property TextBuffer As ITextBuffer
            Get
                Dim textBuffer1 As ITextBuffer = CType(GetValue(TextBufferProperty), ITextBuffer)
                If textBuffer1 Is Nothing AndAlso BindingOperations.GetBinding(Me, TextBufferProperty) Is Nothing Then
                    textBuffer1 = If(BufferFactory Is Nothing,
                            New BufferFactory().CreateTextBuffer("", ContentType),
                            BufferFactory.CreateTextBuffer("", ContentType)
                        )
                End If

                Return textBuffer1
            End Get

            Set(value As ITextBuffer)
                SetValue(TextBufferProperty, value)
                ContentType = value.ContentType
            End Set
        End Property

        <Import>
        Public Property TextStructureNavigatorFactory As ITextStructureNavigatorFactory

        Public ReadOnly Property TextView As IAvalonTextView
            Get
                If _textView Is Nothing Then
                    Initialize()
                End If

                Return _textView
            End Get
        End Property

        Public ReadOnly Property TextViewHost As IAvalonTextViewHost
            Get
                If _textViewHost Is Nothing Then
                    Initialize()
                End If

                Return _textViewHost
            End Get
        End Property

        <Import(GetType(TextViewService))>
        Public Property TextViewServices As IEnumerable(Of ImportInfo(Of Action(Of ITextView), IContentTypeMetadata))

        <Import(GetType(AdornmentProviderFactory))>
        Public Property AdornmentProviderFactories As IEnumerable(Of ImportInfo(Of Func(Of ITextView, IAdornmentProvider), IContentTypeMetadata))

        <Import>
        Public Property UndoHistoryRegistry As IUndoHistoryRegistry

        Public Event PropertyChanged As PropertyChangedEventHandler Implements ComponentModel.INotifyPropertyChanged.PropertyChanged

        Public Sub New()
            MyBase.Background = Brushes.Transparent
            Dim ctlg = Catalog.GlobalCatalog
            If ctlg IsNot Nothing Then
                Dim domain As New ComponentDomain(New CatalogResolver(ctlg))
                domain.AddComponent(Me)
                domain.Bind()
            End If
        End Sub

        Private Shared Sub TextBufferChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            CType(d, CodeEditorControl).Content = Nothing
            CType(d, CodeEditorControl).InvalidateMeasure()
        End Sub

        Public Sub ClearHighlighting()
            If _ContainsHighlights OrElse _ContainsWordHighlights Then
                Dim marker = FindMarkerProvider.GetFindMarkerProvider(TextView)
                marker.ClearAllMarkers()
                _ContainsHighlights = False
                _ContainsWordHighlights = False
            End If
        End Sub

        Public Function HighlightNextMatch(searchText As String, ignoreCase As Boolean) As Boolean
            Dim textSnapshot1 = TextView.TextSnapshot
            Dim __ As Object = String.Empty
            __ = textSnapshot1.LineCount

            Dim line = textSnapshot1.GetLineFromPosition(TextView.Caret.Position.TextInsertionIndex)
            Dim index = TextView.Caret.Position.TextInsertionIndex
            Dim searchHit = GetSearchHit(line, index - line.Start, searchText, ignoreCase)

            If searchHit IsNot Nothing Then
                TextView.Selection.ActiveSpan = searchHit
                TextView.Caret.MoveTo(searchHit.GetEndPoint(textSnapshot1))
                TextView.Caret.EnsureVisible()
                Return True
            End If

            For i = line.LineNumber + 1 To textSnapshot1.LineCount - 1
                Dim line2 = textSnapshot1.GetLineFromLineNumber(i)
                searchHit = GetSearchHit(line2, 0, searchText, ignoreCase)
                If searchHit IsNot Nothing Then
                    TextView.Selection.ActiveSpan = searchHit
                    TextView.Caret.MoveTo(searchHit.GetEndPoint(textSnapshot1))
                    TextView.Caret.EnsureVisible()
                    Return True
                End If
            Next

            For j = 0 To line.LineNumber
                Dim line2 = textSnapshot1.GetLineFromLineNumber(j)
                searchHit = GetSearchHit(line2, 0, searchText, ignoreCase)
                If searchHit IsNot Nothing Then
                    TextView.Selection.ActiveSpan = searchHit
                    TextView.Caret.MoveTo(searchHit.GetEndPoint(textSnapshot1))
                    TextView.Caret.EnsureVisible()
                    Return True
                End If
            Next
            Return False
        End Function

        Public Function GetSearchHit(
                    textLine As ITextSnapshotLine,
                    startIndex As Integer,
                    searchText As String,
                    ignoreCase As Boolean
            ) As TextSpan

            Dim comparisonType = (If(ignoreCase, StringComparison.InvariantCultureIgnoreCase, StringComparison.InvariantCulture))
            Dim text = textLine.GetText()
            Dim index = text.IndexOf(searchText, startIndex, comparisonType)
            If index <> -1 Then
                Return New TextSpan(
                    textLine.TextSnapshot,
                    New Span(index + textLine.Start, searchText.Length),
                    SpanTrackingMode.EdgeInclusive)

            End If

            Return Nothing
        End Function

        Public Sub HighlightMatches(
                    searchText As String,
                    ignoreCase As Boolean
            )

            ClearHighlighting()
            Dim textSnapshot1 = TextView.TextSnapshot
            Dim text = ""
            Dim index = -1
            Dim length = searchText.Length
            Dim comparisonType = (If(ignoreCase, StringComparison.InvariantCultureIgnoreCase, StringComparison.InvariantCulture))
            Dim markers As New List(Of FindMarker)

            For i = 0 To textSnapshot1.LineCount - 1
                Dim line = textSnapshot1.GetLineFromLineNumber(i)
                text = line.GetText()
                For j As Integer = 0 To text.Length - 1
                    index = text.IndexOf(searchText, j, comparisonType)
                    If index = -1 Then Exit For
                    If index + line.Start <> _textView.Selection.ActiveSnapshotSpan.Start Then
                        markers.Add(New FindMarker(
                                  New TextSpan(textSnapshot1, index + line.Start, length, SpanTrackingMode.EdgeInclusive),
                                  Color.FromArgb(96, 255, 255, 0)
                               ))
                        j = index
                    End If
                Next
            Next

            If markers.Count > 0 Then
                Dim provider = FindMarkerProvider.GetFindMarkerProvider(TextView)
                provider.AddFindMarkers(markers)
                _ContainsHighlights = True
            End If
        End Sub


        Public Property WordHighlightingColor As Color = Colors.LightGray

        Public Sub HighlightWords(ParamArray spans() As (Start As Integer, Lenght As Integer))

            If spans Is Nothing OrElse spans.Count = 0 Then Return

            ClearHighlighting()
            Dim textSnapshot1 = TextView.TextSnapshot
            Dim markers As New List(Of FindMarker)

            For Each span In spans
                markers.Add(New FindMarker(
                          New TextSpan(textSnapshot1, span.Start, span.Lenght,
                                       SpanTrackingMode.EdgeInclusive),
                          WordHighlightingColor
                       ))
            Next

            Dim provider = FindMarkerProvider.GetFindMarkerProvider(TextView)
            provider.AddFindMarkers(markers)
            _ContainsWordHighlights = True
        End Sub

        Public Sub SelectAnotherHighlightedWord(moveDown As Boolean)
            If Not _ContainsWordHighlights Then Return
            Dim provider = FindMarkerProvider.GetFindMarkerProvider(TextView)
            Dim marker = provider.GetFindMarker(moveDown)
            If marker Is Nothing Then Return

            Dim span = marker.Span
            TextView.Selection.ActiveSpan = span
            TextView.Caret.MoveTo(span.GetEndPoint(TextView.TextSnapshot))
            TextView.Caret.EnsureVisible()
            TextView.Caret.MoveTo(span.GetStartPoint(TextView.TextSnapshot))
            TextView.Caret.EnsureVisible()
        End Sub

        Private Sub Initialize()
            If MyBase.Content Is Nothing AndAlso TextBuffer IsNot Nothing Then
                _textView = New AvalonTextView(TextBuffer, ClassifierAggregatorProvider, AdornmentAggregatorCache, AdornmentSurfaceManagerFactory, ClassificationFormatMap, ClassificationTypeRegistry, TextViewServices)
                _textViewHost = New AvalonTextViewHost(UndoHistoryRegistry, Orderer, TextView, AvalonTextViewMarginFactories, EditorOperationsProvider, MouseBindingFactories, KeyboardFilters)
                _textViewHost.TextView.Background = MyBase.Background
                _textViewHost.ScaleFactor = _scaleFactor

                If Not IsLineNumberMarginVisible Then
                    _textViewHost.GetTextViewMargin("Avalon Line Number Margin").MarginVisible = False
                End If

                AddHandler _textView.Selection.SelectionChanged, AddressOf OnSelectionChanged
                MyBase.Content = _textViewHost

                If Keyboard.FocusedElement Is Nothing Then
                    _textView.VisualElement.Focus()
                End If
            End If
        End Sub

        Protected Overrides Function MeasureOverride(constraint As Size) As Size
            Initialize()
            Return MyBase.MeasureOverride(constraint)
        End Function

        Private Sub Notify(prop As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(prop))
        End Sub

        Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
            Dim flag As Boolean = (Keyboard.Modifiers And ModifierKeys.Shift) = ModifierKeys.Shift
            If (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then
                If e.Key = Key.Add OrElse (e.Key = Key.OemPlus AndAlso flag) Then
                    ScaleFactor = Math.Round(ScaleFactor * 1.02, 2)
                    e.Handled = True

                ElseIf e.Key = Key.Subtract OrElse e.Key = Key.OemMinus Then
                    ScaleFactor = Math.Round(ScaleFactor / 1.02, 2)
                    e.Handled = True

                End If
            End If

            If e.Key = Key.Escape Then e.Handled = True

            MyBase.OnKeyDown(e)
        End Sub

        Protected Overrides Sub OnPreviewMouseWheel(e As MouseWheelEventArgs)
            If (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then
                Dim n = 1.0 + 0.05 * (CDbl(e.Delta) / 120.0)
                ScaleFactor *= n
                e.Handled = True
            End If
            MyBase.OnMouseWheel(e)
        End Sub

        Private Sub OnSelectionChanged(sender As Object, e As EventArgs)
            If HighlightSearchHits AndAlso _textView.Selection.ActiveSnapshotSpan.Length > 2 Then
                Dim text As String = _textView.Selection.ActiveSnapshotSpan.GetText()
                HighlightMatches(text, ignoreCase:=True)
            Else
                ClearHighlighting()
            End If
        End Sub

    End Class
End Namespace
