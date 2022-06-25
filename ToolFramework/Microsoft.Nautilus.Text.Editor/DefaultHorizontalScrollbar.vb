Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class DefaultHorizontalScrollbar
        Implements IAvalonTextViewMargin, ITextViewMargin

        Public Const Name As String = "Avalon Horizontal Scrollbar"
        Private Const _smallScroll As Double = 12.0
        Private Const _largeScroll As Double = 72.0
        Private _scrollbar As ScrollBar
        Private _textView As IAvalonTextView

        Public ReadOnly Property VisualElement As FrameworkElement Implements IAvalonTextViewMargin.VisualElement
            Get
                Return _scrollbar
            End Get
        End Property

        Public ReadOnly Property MarginSize As Double Implements ITextViewMargin.MarginSize
            Get
                Return _scrollbar.ActualHeight
            End Get
        End Property

        Public Property MarginVisible As Boolean Implements ITextViewMargin.MarginVisible
            Get
                Return _scrollbar.Visibility = Visibility.Visible
            End Get

            Set(value As Boolean)
                If value <> MarginVisible Then
                    If Not value Then
                        UnRegisterEvents()
                        _scrollbar.Visibility = Visibility.Collapsed
                    Else
                        RegisterEvents()
                        _scrollbar.Visibility = Visibility.Visible
                    End If
                End If
            End Set
        End Property

        Public Sub New(textView As IAvalonTextView)
            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            _textView = textView
            _scrollbar = New ScrollBar
            _scrollbar.Orientation = Orientation.Horizontal
            _scrollbar.SetValue(RangeBase.SmallChangeProperty, 12.0)
            _scrollbar.SetValue(RangeBase.LargeChangeProperty, 72.0)
            _scrollbar.Minimum = 0.0
            AddHandler _textView.WordWrapStyleChanged, AddressOf WordWrapChanged
            WordWrapChanged(Nothing, Nothing)
            RegisterEvents()
        End Sub

        Private Sub RegisterEvents()
            AddHandler _textView.LayoutChanged, AddressOf EditorLayoutChanged
            AddHandler _textView.ViewportLeftChanged, AddressOf EditorLayoutChanged
            AddHandler _textView.ViewportWidthChanged, AddressOf EditorLayoutChanged
            AddHandler _scrollbar.Scroll, AddressOf HorizontalScrollBarScrolled
        End Sub

        Private Sub UnRegisterEvents()
            RemoveHandler _textView.LayoutChanged, AddressOf EditorLayoutChanged
            RemoveHandler _textView.ViewportLeftChanged, AddressOf EditorLayoutChanged
            RemoveHandler _textView.ViewportWidthChanged, AddressOf EditorLayoutChanged
            RemoveHandler _scrollbar.Scroll, AddressOf HorizontalScrollBarScrolled
        End Sub

        Friend Sub HorizontalScrollBarScrolled(sender As Object, e As ScrollEventArgs)
            _textView.ViewportLeft = e.NewValue
        End Sub

        Friend Sub EditorLayoutChanged(sender As Object, e As EventArgs)
            _scrollbar.Maximum = Math.Max(_textView.MaxTextWidth + 10.0 - _textView.ViewportWidth, _textView.ViewportLeft)
            _scrollbar.ViewportSize = _textView.ViewportWidth
            _scrollbar.Value = _textView.ViewportLeft
        End Sub

        Friend Sub WordWrapChanged(sender As Object, e As EventArgs)
            MarginVisible = (_textView.WordWrapStyle And WordWrapStyles.WordWrap) <> WordWrapStyles.WordWrap
        End Sub
    End Class
End Namespace
