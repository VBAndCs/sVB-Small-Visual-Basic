Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports System.Timers
Imports System.Windows
Imports System.Windows.Automation.Peers
Imports System.Windows.Automation.Provider
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Interop
Imports System.Windows.Media
Imports System.Windows.Threading
Imports Microsoft.Nautilus.Core
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Classification
Imports Microsoft.Nautilus.Text.Editor.Automation.Implementation
Imports System.Runtime.InteropServices
Imports Microsoft.Windows.Controls
Imports System.Windows.Controls.Primitives

Namespace Microsoft.Nautilus.Text.Editor
    Public NotInheritable Class AvalonTextView
        Inherits ContentControl
        Implements IAvalonTextView, ITextView, IPropertyOwner, IAdornmentSurfaceHost, IValueProvider

        Private Class LineWidthCache

            Private Class LineWidth

                Public Property Width As Double

                Public ReadOnly Property Span As ITextSpan

                Public Sub New(snapshot As ITextSnapshot, span As Span, width As Double)
                    _Span = snapshot.CreateTextSpan(span, SpanTrackingMode.EdgeInclusive)
                    _Width = width
                End Sub

            End Class

            Private _maxWidth As Double
            Private _cachedLinesCount As Integer
            Private _cachedLines As LineWidth() = New LineWidth(49) {}
            Private _cacheValid As Boolean = True

            Public ReadOnly Property MaxWidth As Double
                Get
                    If _maxWidth = Double.MaxValue Then
                        If _cachedLinesCount = 0 Then
                            Return 0.0
                        End If

                        _maxWidth = _cachedLines(0).Width
                        For i As Integer = 1 To _cachedLinesCount - 1
                            Dim lineWidth1 As LineWidth = _cachedLines(i)
                            If lineWidth1.Width > _maxWidth Then
                                _maxWidth = lineWidth1.Width
                            End If
                        Next
                    End If

                    Return _maxWidth
                End Get
            End Property

            Public Sub MarkCacheInvalid()
                _cacheValid = False
            End Sub

            Public Sub AddLine(snapshot As ITextSnapshot, span As Span, width As Double)
                If width > _maxWidth Then
                    _maxWidth = width
                End If

                For i As Integer = 0 To _cachedLinesCount - 1
                    Dim lineWidth = _cachedLines(i)
                    If CInt(CLng(Fix(lineWidth.Span.GetStartPoint(snapshot))) Mod Integer.MaxValue) = span.Start Then
                        If lineWidth.Width = _maxWidth AndAlso width < lineWidth.Width Then
                            _maxWidth = Double.MaxValue
                        End If
                        lineWidth.Width = width
                        Return
                    End If
                Next

                If _cachedLinesCount < 50 Then
                    Dim n = _cachedLinesCount
                    _cachedLines(n) = New LineWidth(snapshot, span, width)
                    _cachedLinesCount = n + 1
                    Return
                End If

                Dim num As Integer = 0
                For j As Integer = 1 To _cachedLinesCount - 1
                    If _cachedLines(j).Width < _cachedLines(num).Width Then
                        num = j
                    End If
                Next

                If width > _cachedLines(num).Width Then
                    _cachedLines(num) = New LineWidth(snapshot, span, width)
                End If
            End Sub

            Public Sub InvalidateTextLine(snapshot As ITextSnapshot, span As Span)
                Dim num As Integer = 0
                While num < _cachedLinesCount
                    Dim lineWidth = _cachedLines(num)

                    If lineWidth.Span.GetSpan(snapshot).OverlapsWith(span) Then

                        If lineWidth.Width = _maxWidth Then
                            _maxWidth = Double.MaxValue
                        End If

                        _cachedLinesCount -= 1
                        _cachedLines(num) = _cachedLines(_cachedLinesCount)
                        _cachedLines(_cachedLinesCount) = Nothing

                    Else
                        num += 1
                    End If

                End While
            End Sub

            Public Sub FlushCacheIfInvalid()
                If _cacheValid Then Return

                _cacheValid = True
                _maxWidth = 0.0

                If _cachedLinesCount > 0 Then
                    For i As Integer = 0 To _cachedLinesCount - 1
                        _cachedLines(i) = Nothing
                    Next
                    _cachedLinesCount = 0
                End If
            End Sub
        End Class

        Public Property Editor As CodeEditorControl
        Private Const WM_IME_STARTCOMPOSITION As Integer = 269
        Private Const WM_IME_ENDCOMPOSITION As Integer = 270
        Private _defaultFormattedTextLines As IFormattedTextLineCollection
        Private _nextTextSnapshot As ITextSnapshot
        Private _classifier As IClassifier
        Private _adornmentProvider As IAdornmentProvider
        Private _visualsFactory As VisualsFactory
        Private _controlHostLayer As Canvas
        Private _baseLayer As Canvas
        Private _contentLayer As TextContentLayer
        Private _selectionLayer As SelectionLayer
        Private _gapBehavior As GapBehaviors = GapBehaviors.GapAtBottomAllowed
        Private _mouseHoverEventLock As New Object
        Friend _mouseHoverTimer As Timer
        Private _lastHoverPosition As Integer? = Nothing
        Private _raiseHoverEvent As Boolean = True
        Private _topAnchorCharacterPosition As Integer
        Private _topAnchorTop As Double
        Private _wordWrapStyle As WordWrapStyles
        Private _textLineVisuals As List(Of TextLineVisual)
        Private _invalidationLock As New Object
        Private _invalidTextLineSpans As New NormalizedSpanCollection
        Private _invalidAdornmentSpans As New NormalizedSpanCollection
        Private _adornmentList As List(Of IAdornment)
        Private _lineWidthCache As LineWidthCache
        Private _viewportLeft As Double
        Private _caret As CaretElement
        Private _adornSurface As IAdornmentSurfaceManager
        Private _layoutChangeStart As Integer = Integer.MaxValue
        Private _layoutChangeEnd As Integer
        Private _immHookSet As Boolean
        Private _imeEnabled As Boolean
        Private _immHook As HwndSourceHook
        Private _layoutNeeded As Boolean = True

        Public ReadOnly Property TextView As IAvalonTextView Implements IAdornmentSurfaceHost.TextView
            Get
                Return Me
            End Get
        End Property

        ReadOnly Property IValueProvider_Value As String Implements IValueProvider.Value
            Get
                Return _TextSnapshot.GetText(New Span(0, _TextSnapshot.Length))
            End Get
        End Property

        ReadOnly Property IValueProvider_IsReadOnly As Boolean = False Implements IValueProvider.IsReadOnly

        Public ReadOnly Property Properties As New PropertyCollection Implements IPropertyOwner.Properties

        Public Shadows Property Background As Brush Implements IAvalonTextView.Background
            Get
                Return _controlHostLayer.Background
            End Get

            Set(value As Brush)
                _controlHostLayer.Background = value
            End Set
        End Property

        Public ReadOnly Property Caret As ITextCaret Implements IAvalonTextView.Caret
            Get
                Return _caret
            End Get
        End Property

        Public ReadOnly Property VisualElement As FrameworkElement Implements IAvalonTextView.VisualElement
            Get
                Return Me
            End Get
        End Property

        Public ReadOnly Property SpanGeometry As ISpanGeometry Implements IAvalonTextView.SpanGeometry
            Get
                Return New DefaultSpanGeometry(Me)
            End Get
        End Property

        Public ReadOnly Property Selection As ITextSelection Implements ITextView.Selection
            Get
                Return _selectionLayer
            End Get
        End Property

        Public ReadOnly Property TextBuffer As ITextBuffer Implements ITextView.TextBuffer

        Public ReadOnly Property TextSnapshot As ITextSnapshot Implements ITextView.TextSnapshot

        Public ReadOnly Property FirstVisibleCharacterPosition As Integer Implements ITextView.FirstVisibleCharacterPosition
            Get
                Return _topAnchorCharacterPosition
            End Get
        End Property

        Public ReadOnly Property FormattedTextLines As IFormattedTextLineCollection Implements ITextView.FormattedTextLines
            Get
                Return _defaultFormattedTextLines
            End Get
        End Property

        Dim _viewScroller As DefaultViewScroller

        Public ReadOnly Property ViewScroller As IViewScroller Implements ITextView.ViewScroller
            Get
                If _viewScroller Is Nothing Then
                    _viewScroller = New DefaultViewScroller(Me)
                End If

                Return _viewScroller
            End Get
        End Property

        Public ReadOnly Property MaxTextWidth As Double Implements ITextView.MaxTextWidth

        Public Property ViewportLeft As Double Implements ITextView.ViewportLeft
            Get
                Return _viewportLeft
            End Get

            Set(value As Double)
                Dim num As Double
                If (_wordWrapStyle And WordWrapStyles.WordWrap) > 0 Then
                    num = 0.0
                Else
                    If _layoutNeeded Then PerformLayout()
                    num = Math.Max(0.0, Math.Min(value, MaxTextWidth + 10.0 - ViewportWidth))
                End If

                If _viewportLeft <> num Then
                    _viewportLeft = num
                    Canvas.SetLeft(_baseLayer, 0.0 - _viewportLeft)
                    RaiseEvent ViewportLeftChanged(Me, New EventArgs)
                    If _imeEnabled Then PositionImmCompositionWindow()
                End If
            End Set
        End Property

        Public ReadOnly Property ViewportTop As Double = 0.0 Implements ITextView.ViewportTop

        Public ReadOnly Property ViewportRight As Double Implements ITextView.ViewportRight
            Get
                Return _viewportLeft + ViewportWidth
            End Get
        End Property

        Public ReadOnly Property ViewportBottom As Double Implements ITextView.ViewportBottom
            Get
                Return _ViewportTop + ViewportHeight
            End Get
        End Property

        Public ReadOnly Property ViewportHeight As Double Implements ITextView.ViewportHeight
            Get
                If Not Double.IsNaN(MyBase.ActualHeight) Then
                    Return MyBase.ActualHeight
                End If

                Return 120.0
            End Get
        End Property

        Public ReadOnly Property ViewportWidth As Double Implements ITextView.ViewportWidth
            Get
                If Not Double.IsNaN(MyBase.ActualWidth) Then
                    Return MyBase.ActualWidth
                End If

                Return 240.0
            End Get
        End Property

        Public Property WordWrapStyle As WordWrapStyles Implements ITextView.WordWrapStyle
            Get
                Return _wordWrapStyle
            End Get

            Set(value As WordWrapStyles)
                If (CUInt(value) And &HFFFFFFF8UI) <> 0 Then
                    Throw New ArgumentOutOfRangeException("value")
                End If

                If value <> _wordWrapStyle Then
                    _wordWrapStyle = value
                    If (_wordWrapStyle And WordWrapStyles.WordWrap) = WordWrapStyles.WordWrap Then
                        ViewportLeft = 0.0
                    End If

                    InvalidateAllLines()
                    RaiseEvent WordWrapStyleChanged(Me, New EventArgs)
                End If
            End Set
        End Property

        Private _mouseHover As EventHandler(Of MouseHoverEventArgs)

        Public Event LayoutChanged As EventHandler(Of TextViewLayoutChangedEventArgs) Implements ITextView.LayoutChanged

        Public Event ViewportLeftChanged As EventHandler Implements ITextView.ViewportLeftChanged

        Public Event ViewportHeightChanged As EventHandler Implements ITextView.ViewportHeightChanged

        Public Event ViewportWidthChanged As EventHandler Implements ITextView.ViewportWidthChanged

        Public Event WordWrapStyleChanged As EventHandler Implements ITextView.WordWrapStyleChanged

        Public Custom Event MouseHover As EventHandler(Of MouseHoverEventArgs) Implements ITextView.MouseHover
            AddHandler(value As EventHandler(Of MouseHoverEventArgs))
                SyncLock _mouseHoverEventLock
                    [Delegate].Combine(_mouseHover, value)
                    EnableMouseHover()
                End SyncLock
            End AddHandler

            RemoveHandler(value As EventHandler(Of MouseHoverEventArgs))
                SyncLock _mouseHoverEventLock
                    [Delegate].Remove(_mouseHover, value)
                    If _mouseHover Is Nothing Then DisableMouseHover()
                End SyncLock
            End RemoveHandler

            RaiseEvent(sender As Object, e As MouseHoverEventArgs)
                _mouseHover?.Invoke(sender, e)
            End RaiseEvent
        End Event

        Public Sub New(
                      textBuffer As ITextBuffer,
                      classifierAggregatorProvider As IClassifierAggregatorProvider,
                      adornmentAggregatorCache As IAdornmentAggregatorCache,
                      adornmentSurfaceManagerFactory As IAdornmentSurfaceManagerFactory,
                      classificationFormatMap As IClassificationFormatMap,
                      classificationTypeRegistry As IClassificationTypeRegistry,
                      textViewServices As IEnumerable(Of ImportInfo(Of Action(Of ITextView), IContentTypeMetadata))
                )

            InputMethod.SetIsInputMethodSuspended(Me, value:=True)
            _TextBuffer = textBuffer
            _nextTextSnapshot = _TextBuffer.CurrentSnapshot
            _TextSnapshot = _nextTextSnapshot
            _topAnchorCharacterPosition = 0
            _topAnchorTop = 0.0
            _lineWidthCache = New LineWidthCache
            _adornSurface = adornmentSurfaceManagerFactory.CreateAdornmentSurfaceManager(Me)
            _immHook = AddressOf WndProc
            _textLineVisuals = New List(Of TextLineVisual)
            _defaultFormattedTextLines = New DefaultFormattedTextLineCollection(Of TextLineVisual)(Me, _textLineVisuals)
            _classifier = classifierAggregatorProvider.GetClassifier(_TextBuffer)
            _adornmentProvider = adornmentAggregatorCache.GetAdornmentProvider(Me)

            AddHandler _classifier.ClassificationChanged, AddressOf OnClassificationChanged
            AddHandler _adornmentProvider.AdornmentsChanged, AddressOf OnAdornmentsChanged

            _visualsFactory = New VisualsFactory(Me, _classifier, classificationFormatMap, classificationTypeRegistry)
            InitializeLayers()
            _adornmentList = New List(Of IAdornment)
            _caret = New CaretElement(Me)
            _baseLayer.Children.Add(_caret)
            _caret.Visibility = Visibility.Hidden

            AddHandler _TextBuffer.Changed, AddressOf OnTextBufferChanged
            AddHandler MyBase.SizeChanged, AddressOf OnSizeChanged
            AddHandler MyBase.GotKeyboardFocus, AddressOf OnGotFocus
            AddHandler MyBase.LostKeyboardFocus, AddressOf OnLostFocus

            For Each service In textViewServices
                For Each contentType In service.Metadata.ContentTypes
                    If ContentTypeHelper.IsOfType(_TextBuffer.ContentType, contentType) Then
                        service.GetBoundValue()(Me)
                        Exit For
                    End If
                Next
            Next
        End Sub

        Public Sub DisplayTextLineContainingCharacter(position As Integer, verticalDistance As Double, relativeTo As ViewRelativePosition) Implements ITextView.DisplayTextLineContainingCharacter
            If position < 0 OrElse position > _TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            If relativeTo <> 0 AndAlso relativeTo <> ViewRelativePosition.Bottom Then
                Throw New ArgumentOutOfRangeException("relativeTo")
            End If

            PerformLayout(position, verticalDistance, relativeTo, _gapBehavior, ignoreViewHeight:=False)
        End Sub

        Public Sub DisplayAllContent() Implements ITextView.DisplayAllContent
            PerformLayout(0, 0.0, ViewRelativePosition.Top, GapBehaviors.GapAtTopAllowed Or GapBehaviors.GapAtBottomAllowed, ignoreViewHeight:=True)
        End Sub

        Public Sub Invalidate() Implements IAvalonTextView.Invalidate
            InvalidateAllLines()
        End Sub

        Public Function GetTextElementSpan(position As Integer) As Span Implements ITextView.GetTextElementSpan
            If position < 0 OrElse position > _TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            Dim textLineVisual = TryCast(FormattedTextLines.GetTextLineContainingPosition(position), TextLineVisual)
            If textLineVisual Is Nothing Then
                Dim line = _TextSnapshot.GetLineFromPosition(position)
                Dim lineVisuals = LayoutTextLinesForPositioningCaret(line)

                For Each lineVisual In lineVisuals
                    If lineVisual.ContainsPosition(position) Then
                        textLineVisual = lineVisual
                        Exit For
                    End If
                Next

            End If

            Return If(textLineVisual?.GetAvalonTextElementSpan(position), New Span(position, Math.Min(_TextSnapshot.Length - position, 1)))
        End Function

        Public Sub AddAdornmentSurface(surface As IAdornmentSurface) Implements IAdornmentSurfaceHost.AddAdornmentSurface
            If surface Is Nothing Then
                Throw New ArgumentNullException("adornmentSurface")
            End If

            Dim panel = surface.SurfacePanel
            Dim children = _baseLayer.Children

            Select Case surface.SurfacePosition
                Case SurfacePosition.Bottommost
                    children.Insert(0, panel)

                Case SurfacePosition.BelowSelection
                    children.Insert(children.IndexOf(_selectionLayer), panel)

                Case SurfacePosition.AboveSelection
                    children.Insert(children.IndexOf(_selectionLayer) + 1, panel)

                Case SurfacePosition.BelowText
                    children.Insert(children.IndexOf(_contentLayer), panel)

                Case SurfacePosition.AboveText
                    children.Insert(children.IndexOf(_contentLayer) + 1, panel)

                Case Else
                    children.Add(panel)

            End Select

            panel.Width = _baseLayer.ActualWidth
            panel.Height = _baseLayer.ActualHeight

            AddHandler _baseLayer.SizeChanged,
                Sub()
                    panel.Width = _baseLayer.ActualWidth
                    panel.Height = _baseLayer.ActualHeight
                End Sub
        End Sub

        Private Overloads Sub SetValue(text As String) Implements IValueProvider.SetValue
            Dim textEdit As ITextEdit

            Do
                textEdit = _TextBuffer.CreateEdit()
                Dim span As New Span(0, textEdit.Snapshot.Length)
                If Not textEdit.CanDeleteOrReplace(span) Then
                    Exit Do
                End If

                textEdit.Replace(span, text)
            Loop While textEdit.Apply() Is Nothing
        End Sub

        Protected Overrides Function ArrangeOverride(arrangeBounds As Size) As Size
            If _layoutNeeded Then
                PerformLayout()
            End If
            Return MyBase.ArrangeOverride(arrangeBounds)
        End Function

        Protected Overrides Function MeasureOverride(availableSize As Size) As Size
            If availableSize.Width = Double.PositiveInfinity OrElse availableSize.Height = Double.PositiveInfinity Then
                Return New Size(240.0, 120.0)
            End If
            Return availableSize
        End Function

        Protected Overrides Function OnCreateAutomationPeer() As AutomationPeer
            Return New TextEditorAutomationPeer(Me)
        End Function

        Protected Overrides Sub OnGotKeyboardFocus(e As KeyboardFocusChangedEventArgs)
            If Not _immHookSet Then
                AvalonHelper.GetHwndSource(Me)?.AddHook(_immHook)
                _immHookSet = True
            End If
            AvalonHelper.EnableImmComposition(Me)
            MyBase.OnGotKeyboardFocus(e)
        End Sub

        Protected Overrides Sub OnLostKeyboardFocus(e As KeyboardFocusChangedEventArgs)
            AvalonHelper.DisableImmComposition(Me)
            If _immHookSet Then
                AvalonHelper.GetHwndSource(Me)?.RemoveHook(_immHook)
                _immHookSet = False
            End If
            MyBase.OnLostKeyboardFocus(e)
        End Sub

        Private Function WndProc(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr, ByRef handled As Boolean) As IntPtr
            Select Case msg
                Case 269
                    _imeEnabled = True
                    PositionImmCompositionWindow()
                    _caret.EnsureVisible()
                    _caret.Visibility = Visibility.Hidden

                Case 270
                    _imeEnabled = False
                    _caret.Visibility = Visibility.Visible
            End Select
            Return IntPtr.Zero
        End Function

        Private Sub OnTextBufferChanged(sender As Object, e As TextChangedEventArgs)
            Dim normalizedSpans As NormalizedSpanCollection
            If e.Changes Is Nothing Then
                normalizedSpans = New NormalizedSpanCollection()
            Else
                Dim spans As New List(Of Span)
                Dim delta As Integer = 0
                For Each change In e.Before.Version.Changes
                    Dim s = Span.FromBounds(change.Position - delta, change.OldEnd - delta)
                    spans.Add(GetInvalidationSpan(s))
                    delta += change.Delta
                Next
                normalizedSpans = New NormalizedSpanCollection(spans)
            End If

            If e.After.Version > _TextSnapshot.Version Then
                SyncLock _invalidationLock
                    _nextTextSnapshot = e.After
                    InvalidateLines(normalizedSpans)
                End SyncLock
            End If

            PerformLayout()
        End Sub

        Private Sub OnSizeChanged(sender As Object, e As SizeChangedEventArgs)
            If _controlHostLayer.Height <> MyBase.ActualHeight Then
                _layoutNeeded = True
                InvalidateArrange()
                _controlHostLayer.Height = MyBase.ActualHeight
                RaiseEvent ViewportHeightChanged(Me, New EventArgs)
            End If

            If _controlHostLayer.Width <> MyBase.ActualWidth Then
                _controlHostLayer.Width = MyBase.ActualWidth
                If (WordWrapStyle And WordWrapStyles.WordWrap) = WordWrapStyles.WordWrap Then
                    InvalidateAllLines()
                End If

                RaiseEvent ViewportWidthChanged(Me, New EventArgs)
            End If
        End Sub

        Private Overloads Sub OnGotFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
            _caret.Visibility = Visibility.Visible
            EnableMouseHover()
        End Sub

        Private Overloads Sub OnLostFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
            _caret.Visibility = Visibility.Hidden
            DisableMouseHover()
        End Sub

        Private Sub OnClassificationChanged(sender As Object, e As ClassificationChangedEventArgs)
            Dim invalidationSpan As Span = GetInvalidationSpan(e.ChangeSpan.GetSpan(_TextSnapshot))
            SyncLock _invalidationLock
                InvalidateLines(invalidationSpan)
            End SyncLock
        End Sub

        Private Sub OnAdornmentsChanged(sender As Object, e As AdornmentsChangedEventArgs)
            Dim invalidationSpan As Span = GetInvalidationSpan(e.ChangeSpan.GetSpan(_TextSnapshot))
            SyncLock _invalidationLock
                InvalidateLines(invalidationSpan)
                InvalidateAdornments(invalidationSpan)
            End SyncLock
        End Sub

        Private Function GetInvalidationSpan(span As Span) As Span
            If (WordWrapStyle And WordWrapStyles.WordWrap) <> WordWrapStyles.WordWrap Then
                Return span
            Else
                Return Span.FromBounds(
                    _TextSnapshot.GetLineFromPosition(span.Start).Start,
                    _TextSnapshot.GetLineFromPosition(span.End).EndIncludingLineBreak
                )
            End If
        End Function

        Protected Overrides Sub OnPreviewMouseLeftButtonDown(e As MouseButtonEventArgs)
            MyBase.OnMouseLeftButtonDown(e)
            If Not MyBase.IsKeyboardFocusWithin Then
                Focus()
            End If
        End Sub

        Private Sub OnHoverTimer(sender As Object, e As ElapsedEventArgs)
            MyBase.Dispatcher.Invoke(DispatcherPriority.Normal, CType(
                    Function() As Object
                        HoverAtPoint(Mouse.PrimaryDevice.GetPosition(Me))
                        Return Nothing
                    End Function, DispatcherOperationCallback), Nothing)
        End Sub

        Private Sub InitializeLayers()
            ClipToBounds = True
            Focusable = True
            FocusVisualStyle = Nothing
            Cursor = Cursors.IBeam
            _baseLayer = New Canvas
            _baseLayer.ClipToBounds = False
            _selectionLayer = New SelectionLayer(Me)
            _contentLayer = New TextContentLayer
            _baseLayer.Children.Add(_selectionLayer)
            _baseLayer.Children.Add(_contentLayer)
            _controlHostLayer = New Canvas
            _controlHostLayer.Background = SystemColors.WindowBrush
            MyBase.Content = _controlHostLayer
            _controlHostLayer.Children.Add(_baseLayer)
            AddHandler _controlHostLayer.SizeChanged,
                Sub()
                    Dim baseLayer As Canvas = _baseLayer
                    Dim selectionLayer1 As SelectionLayer = _selectionLayer
                    Dim num As Double = _controlHostLayer.Width
                    _contentLayer.Width = num
                    baseLayer.Width = num
                    selectionLayer1.Width = num

                    Dim baseLayer2 As Canvas = _baseLayer
                    Dim selectionLayer2 As SelectionLayer = _selectionLayer
                    Dim num5 As Double = _controlHostLayer.Height
                    _contentLayer.Height = num5
                    selectionLayer2.Height = num5
                    baseLayer2.Height = num5
                End Sub
        End Sub

        Friend Sub PerformLayout()
            Dim snapshot = New SnapshotPoint(_TextSnapshot, _topAnchorCharacterPosition)
            Dim anchorPos = snapshot.TranslateTo(_nextTextSnapshot, TrackingMode.Negative)
            PerformLayout(anchorPos, _topAnchorTop, ViewRelativePosition.Top, _gapBehavior, ignoreViewHeight:=False)
        End Sub

        Private Sub PerformLayout(anchorPosition As Integer, verticalDistance As Double, relativeTo As ViewRelativePosition, gapBehavior As GapBehaviors, ignoreViewHeight As Boolean)
            _layoutNeeded = False
            _layoutChangeStart = Integer.MaxValue
            _layoutChangeEnd = 0
            Dim curSnapshot = _TextSnapshot

            SyncLock _invalidationLock
                _lineWidthCache.FlushCacheIfInvalid()
                Dim visuals As New List(Of TextLineVisual)
                Dim i As Integer = 0

                For Each span In _invalidTextLineSpans
                    _lineWidthCache.InvalidateTextLine(_TextSnapshot, span)

                    While i < _textLineVisuals.Count
                        Dim lineVisual = _textLineVisuals(i)
                        If lineVisual.LineSpan.End >= span.Start Then Exit While
                        visuals.Add(lineVisual)
                        i += 1
                    End While

                    While i < _textLineVisuals.Count
                        Dim lineVisual = _textLineVisuals(i)
                        If lineVisual.LineSpan.Start > span.End Then Exit While
                        DisposeTextLineVisual(lineVisual)
                        i += 1
                    End While
                Next

                While i < _textLineVisuals.Count
                    visuals.Add(_textLineVisuals(i))
                    i += 1
                End While

                _textLineVisuals.Clear()
                _textLineVisuals.AddRange(visuals)
                _invalidTextLineSpans = New NormalizedSpanCollection
                Dim adornments As New List(Of IAdornment)

                For Each adornment In _adornmentList
                    Dim spans As New NormalizedSpanCollection(adornment.Span.GetSpan(_TextSnapshot))
                    If _invalidAdornmentSpans.Intersects(spans) Then
                        _adornSurface.RemoveAdornment(adornment)
                    Else
                        adornments.Add(adornment)
                    End If
                Next

                _adornmentList = adornments
                _invalidAdornmentSpans = New NormalizedSpanCollection
                _TextSnapshot = _nextTextSnapshot
            End SyncLock

            InnerPerformLayout(anchorPosition, verticalDistance, relativeTo, gapBehavior, ignoreViewHeight)

            _MaxTextWidth = _lineWidthCache.MaxWidth
            If _imeEnabled Then PositionImmCompositionWindow()

            If LayoutChangedEvent IsNot Nothing Then
                Dim changeSpan As Span = Nothing
                If _layoutChangeStart <= _layoutChangeEnd Then
                    changeSpan = New Span(_layoutChangeStart, _layoutChangeEnd - _layoutChangeStart)
                End If
                RaiseEvent LayoutChanged(Me, New TextViewLayoutChangedEventArgs(changeSpan, curSnapshot, _TextSnapshot))
            End If
        End Sub

        Private Sub InnerPerformLayout(anchorPosition As Integer, verticalDistance As Double, relativeTo As ViewRelativePosition, gapBehavior As GapBehaviors, ignoreViewHeight As Boolean)
            Dim line = _TextSnapshot.GetLineFromPosition(anchorPosition)
            Dim layoutStart As TextLineVisual = Nothing
            Dim layoutEnd As TextLineVisual = Nothing
            Dim topY As Double

            DoAnchorLayout(line, anchorPosition, verticalDistance, relativeTo, layoutStart, layoutEnd, topY)
            layoutStart = If(DoLayoutUp(line, topY), layoutStart)

            If topY > 0.0 Then
                topY = If(
                        (gapBehavior And GapBehaviors.GapAtTopAllowed) <> 0,
                        Math.Min(topY, ViewportHeight - layoutStart.Height),
                        0.0
                    )
            End If

            Dim startIndex = _textLineVisuals.IndexOf(layoutStart)
            Dim index = startIndex - 1
            Dim lineVisual As TextLineVisual

            Do
                index += 1
                lineVisual = _textLineVisuals(index)

                If lineVisual.Top <> topY Then
                    InvalidateLayoutOverSpan(lineVisual.LineSpan)
                End If

                lineVisual.Top = topY
                topY += lineVisual.Height
            Loop While lineVisual IsNot layoutEnd

            Dim i = index + DoLayoutDown(line, topY, ignoreViewHeight)
            If _textLineVisuals(i).Top < 0.0 Then
                InnerPerformLayout(_textLineVisuals(i).LineSpan.Start, 0.0, ViewRelativePosition.Top, GapBehaviors.GapAtTopAllowed Or GapBehaviors.GapAtBottomAllowed, ignoreViewHeight)
                Return
            End If

            If (gapBehavior And GapBehaviors.GapAtBottomAllowed) <> GapBehaviors.GapAtBottomAllowed AndAlso _textLineVisuals(i).Bottom < ViewportHeight Then
                InnerPerformLayout(_textLineVisuals(i).LineSpan.Start, 0.0, ViewRelativePosition.Bottom, gapBehavior Or GapBehaviors.GapAtBottomAllowed, ignoreViewHeight)
                Return
            End If

            RemoveTextLineRange(i + 1, _textLineVisuals.Count)
            RemoveTextLineRange(0, startIndex)

            i = 0
            While i < _textLineVisuals.Count - 1 AndAlso _textLineVisuals(i).Top < 0.0
                i += 1
            End While

            _topAnchorCharacterPosition = _textLineVisuals(i).LineSpan.Start
            _topAnchorTop = _textLineVisuals(i).Top
            TrimAdornments(_textLineVisuals(0).LineSpan.Start, _textLineVisuals(_textLineVisuals.Count - 1).LineSpan.End)
        End Sub

        Private Function DoAnchorLayout(line As ITextSnapshotLine, anchorCharacter As Integer, anchorVerticalDistance As Double, relativeTo As ViewRelativePosition, <Out> ByRef layoutStart As TextLineVisual, <Out> ByRef layoutEnd As TextLineVisual, <Out> ByRef topY As Double) As TextLineVisual
            Dim startIndex, endIndex, index As Integer
            LayoutOneLine(line, anchorVerticalDistance, positionExistingTextLines:=False, startIndex, endIndex)
            layoutStart = _textLineVisuals(startIndex)
            layoutEnd = _textLineVisuals(endIndex)
            _defaultFormattedTextLines.FindTextLineIndexContainingPosition(anchorCharacter, index)

            Dim textLineVisual = _textLineVisuals(index)
            topY = If(relativeTo = ViewRelativePosition.Bottom,
                            ViewportHeight - anchorVerticalDistance - textLineVisual.Height,
                            anchorVerticalDistance)

            For i = index - 1 To startIndex Step -1
                topY -= _textLineVisuals(i).Height
            Next

            Return textLineVisual
        End Function

        Private Function DoLayoutUp(line As ITextSnapshotLine, ByRef topY As Double) As TextLineVisual
            Dim result As TextLineVisual = Nothing
            While line.Start > 0
                line = _TextSnapshot.GetLineFromPosition(line.Start - 1)
                Dim startIndex As Integer
                Dim endIndex As Integer
                Dim num As Double = LayoutOneLine(line, 0.0, positionExistingTextLines:=False, startIndex, endIndex)
                result = _textLineVisuals(startIndex)
                topY -= num
                If topY < 0.0 Then
                    Exit While
                End If
            End While
            Return result
        End Function

        Private Function DoLayoutDown(line As ITextSnapshotLine, bottomY As Double, ignoreViewHeight As Boolean) As Integer
            Dim num As Integer = 0
            While line.End < _TextSnapshot.Length
                line = _TextSnapshot.GetLineFromPosition(line.EndIncludingLineBreak)
                Dim startIndex As Integer
                Dim endIndex As Integer
                Dim num2 = LayoutOneLine(line, bottomY, positionExistingTextLines:=True, startIndex, endIndex)
                num += endIndex - startIndex + 1
                bottomY += num2
                If bottomY > ViewportHeight AndAlso Not ignoreViewHeight Then
                    Exit While
                End If
            End While
            Return num
        End Function

        Private Function LayoutOneLine(line As ITextSnapshotLine, verticalOffset As Double, positionExistingTextLines As Boolean, <Out> ByRef startIndex As Integer, <Out> ByRef endIndex As Integer) As Double
            Dim lineEnd = line.EndIncludingLineBreak
            Dim pos = line.Start
            Dim indent = 0.0
            Dim WordWrapOrAutoIndent = WordWrapStyles.WordWrap Or WordWrapStyles.AutoIndent
            Dim w1 = (If((WordWrapStyle And WordWrapOrAutoIndent) = WordWrapOrAutoIndent, (ViewportWidth * 0.3), 0.0))
            Dim w2 = (If((WordWrapStyle And WordWrapStyles.WordWrap) > 0, ViewportWidth, 0.0))
            Dim h = 0.0

            endIndex = -1
            startIndex = -1

            While _defaultFormattedTextLines.FindTextLineIndexContainingPosition(pos, endIndex)
                Dim lineVisual = _textLineVisuals(endIndex)
                If startIndex = -1 Then
                    startIndex = endIndex
                    If w1 > 0.0 Then
                        indent = lineVisual.GetIndentation()
                        If indent > w1 Then indent = w1
                        w2 -= indent
                        w1 = 0.0
                    End If
                End If

                If positionExistingTextLines AndAlso lineVisual.Top <> verticalOffset Then
                    lineVisual.Top = verticalOffset
                    InvalidateLayoutOverSpan(lineVisual.LineSpan)
                End If

                verticalOffset += lineVisual.Height
                h += lineVisual.Height
                pos = lineVisual.LineSpan.End
                If pos >= lineEnd Then Return h
            End While

            Dim visuals = CreateLineVisuals(line, pos, indent, verticalOffset, w1, w2)
            endIndex += 1
            _textLineVisuals.InsertRange(endIndex, visuals)
            If startIndex = -1 Then startIndex = endIndex
            endIndex += visuals.Count - 1

            For Each visual In visuals
                _lineWidthCache.AddLine(_TextSnapshot, visual.LineSpan, visual.Right)
                _contentLayer.Children.Add(visual)
                verticalOffset += visual.Height
                h += visual.Height
            Next
            Return h
        End Function

        Private Sub InvalidateLayoutOverSpan(span1 As Span)
            If span1.Start < _layoutChangeStart Then
                _layoutChangeStart = span1.Start
            End If

            If span1.End > _layoutChangeEnd Then
                _layoutChangeEnd = span1.End
            End If
        End Sub

        Private Sub RemoveTextLineRange(start1 As Integer, [end] As Integer)
            If start1 < [end] Then
                For i As Integer = start1 To [end] - 1
                    Dim lineVisual As TextLineVisual = _textLineVisuals(i)
                    DisposeTextLineVisual(lineVisual)
                Next
                _textLineVisuals.RemoveRange(start1, [end] - start1)
            End If
        End Sub

        Private Sub DisposeTextLineVisual(lineVisual As TextLineVisual)
            _contentLayer.Children.Remove(lineVisual)
            lineVisual.Dispose()
        End Sub

        Private Sub InvalidateLines(span As Span)
            InvalidateLines(New NormalizedSpanCollection(span))
        End Sub

        Private Sub InvalidateLines(span1 As NormalizedSpanCollection)
            If _textLineVisuals.Count > 0 Then
                _layoutNeeded = True
                InvalidateArrange()
                _invalidTextLineSpans = NormalizedSpanCollection.Union(_invalidTextLineSpans, span1)
            End If
        End Sub

        Private Sub InvalidateAllLines()
            SyncLock _invalidationLock
                _layoutNeeded = True
                InvalidateArrange()
                _invalidTextLineSpans = New NormalizedSpanCollection(New Span(0, _TextSnapshot.Length))
                _lineWidthCache.MarkCacheInvalid()
            End SyncLock
        End Sub

        Private Function CreateLineVisuals(
                            line As ITextSnapshotLine,
                            position As Integer,
                            horizontalOffset As Double,
                            verticalOffset As Double,
                            maxIndent As Double,
                            wrapWidth As Double
                    ) As IList(Of TextLineVisual)

            Dim lineVisuals = _visualsFactory.CreateLineVisuals(line, position, horizontalOffset, verticalOffset, maxIndent, wrapWidth, Nothing)
            Dim adornments = _adornmentProvider.GetAdornments(New SnapshotSpan(_TextSnapshot, New Span(position, line.EndIncludingLineBreak - position)))

            For Each item As IAdornment In adornments
                If Not _adornmentList.Contains(item) Then
                    _adornmentList.Add(item)
                    _adornSurface.AddAdornment(item)
                End If
            Next

            Dim spaceNegotiations As New List(Of SpaceNegotiation)
            For Each lineVisual In lineVisuals
                For Each negotiation In _adornSurface.GetSpaceNegotiations(lineVisual)
                    spaceNegotiations.Add(negotiation)
                Next
            Next

            If spaceNegotiations.Count > 0 Then
                lineVisuals = _visualsFactory.CreateLineVisuals(line, position, horizontalOffset, verticalOffset, maxIndent, wrapWidth, spaceNegotiations)
            End If

            InvalidateLayoutOverSpan(New Span(position, line.EndIncludingLineBreak - position))
            Return lineVisuals
        End Function

        Private Sub InvalidateAdornments(span1 As Span)
            If _textLineVisuals.Count > 0 AndAlso span1.Length > 0 Then
                _layoutNeeded = True
                InvalidateArrange()
                _invalidAdornmentSpans = NormalizedSpanCollection.Union(_invalidAdornmentSpans, New NormalizedSpanCollection(span1))
            End If
        End Sub

        Private Sub TrimAdornments(start1 As Integer, [end] As Integer)
            Dim list1 As New List(Of IAdornment)
            For num As Integer = _adornmentList.Count - 1 To 0 Step -1
                Dim adornment As IAdornment = _adornmentList(num)
                Dim span1 As Span = adornment.Span.GetSpan(_TextSnapshot)
                If span1.End >= start1 AndAlso span1.Start <= [end] Then
                    list1.Add(adornment)
                Else
                    _adornSurface.RemoveAdornment(adornment)
                End If
            Next
            _adornmentList = list1
        End Sub

        Private Sub PositionImmCompositionWindow()
            Dim textLineContainingPosition As ITextLine = _defaultFormattedTextLines.GetTextLineContainingPosition(_caret.Position.TextInsertionIndex)
            If textLineContainingPosition Is Nothing Then
                Return
            End If

            Dim immContext As IntPtr = AvalonHelper.GetImmContext(Me)
            If immContext <> IntPtr.Zero Then
                Dim characterBounds As TextBounds = textLineContainingPosition.GetCharacterBounds(_caret.Position.CharacterIndex)
                Dim num As Double = characterBounds.Left
                If _caret.Position.Placement = CaretPlacement.TrailingEdgeOfCharacter Then
                    num = characterBounds.Right
                End If
                num -= _viewportLeft
                Dim num2 As Double = textLineContainingPosition.Height
                Dim rootVisual As Visual = AvalonHelper.GetRootVisual(Me)
                If rootVisual IsNot Nothing Then
                    Dim generalTransform1 As GeneralTransform = TransformToAncestor(rootVisual)
                    Dim point1 As Point = generalTransform1.Transform(New Point(num, characterBounds.Top))
                    num2 = generalTransform1.Transform(New Point(num, characterBounds.Bottom)).Y - point1.Y
                End If

                AvalonHelper.SetImmFontHeight(immContext, CInt(CLng(Fix(num2)) Mod Integer.MaxValue))
                AvalonHelper.SetCompositionWindowPosition(immContext, New Point(num, characterBounds.Top), Me)
                AvalonHelper.ReleaseImmContext(Me, immContext)
            End If
        End Sub

        Private Function LayoutTextLinesForPositioningCaret(line As ITextSnapshotLine) As IList(Of TextLineVisual)
            Return CreateLineVisuals(line, line.Start, 0.0, -100.0, 0.0, 0.0)
        End Function

        Private Sub EnableMouseHover()
            If _mouseHover IsNot Nothing AndAlso MyBase.IsKeyboardFocusWithin Then
                If _mouseHoverTimer Is Nothing Then
                    _mouseHoverTimer = New Timer(500.0)
                    AddHandler _mouseHoverTimer.Elapsed, AddressOf OnHoverTimer
                End If
                _mouseHoverTimer.Enabled = True
            End If
        End Sub

        Friend Overloads Sub OnMouseMove(sender As Object, e As MouseEventArgs)
            _mouseHoverTimer.Enabled = True
            RemoveHandler MyBase.MouseMove, AddressOf OnMouseMove
        End Sub

        Private Sub DisableMouseHover()
            If _mouseHoverTimer IsNot Nothing Then
                RemoveHandler MyBase.MouseMove, AddressOf OnMouseMove
                _mouseHoverTimer.Dispose()
                _mouseHoverTimer = Nothing
            End If
        End Sub

        Friend Sub HoverAtPoint(pt As Point)
            If _mouseHoverTimer Is Nothing OrElse Not _mouseHoverTimer.Enabled Then
                Return
            End If
            Dim num As Integer? = Nothing
            Dim point1 As New Point(pt.X + ViewportLeft, pt.Y)
            If point1.X >= ViewportLeft AndAlso point1.X < ViewportLeft + ViewportWidth AndAlso point1.Y >= 0.0 AndAlso point1.Y < ViewportHeight Then
                Dim textLineContainingYCoordinate As ITextLine = _defaultFormattedTextLines.GetTextLineContainingYCoordinate(point1.Y)
                If textLineContainingYCoordinate IsNot Nothing Then
                    num = textLineContainingYCoordinate.GetPositionFromXCoordinate(point1.X)
                End If
            End If

            If num <> _lastHoverPosition Then
                _lastHoverPosition = num
                _raiseHoverEvent = True
            ElseIf _raiseHoverEvent AndAlso num.HasValue Then
                _raiseHoverEvent = False
                _mouseHoverTimer.Enabled = False
                AddHandler MyBase.MouseMove, AddressOf OnMouseMove
                Me._mouseHover?.Invoke(Me, New MouseHoverEventArgs(Me, num.Value))
            End If
        End Sub

        Public Event ScrollChaged(senmder As Object, e As ScrollEventArgs) Implements IAvalonTextView.ScrollChaged

        Public Sub OnScrollChanged(e As ScrollEventArgs) Implements IAvalonTextView.OnScrollChanged
            RaiseEvent ScrollChaged(Me, e)
        End Sub
    End Class
End Namespace
