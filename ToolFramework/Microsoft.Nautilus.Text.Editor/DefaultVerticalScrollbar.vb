Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class DefaultVerticalScrollbar
        Implements IAvalonTextViewMargin, ITextViewMargin

        Friend NotInheritable Class DefaultScrollMap
            Private _textView As ITextView

            Public ReadOnly Property ViewportSize As Double
                Get
                    Dim num As Double = _textView.ViewportHeight / 14.0
                    If WordWrapOn Then
                        Return num * (_textView.ViewportWidth / 20.0)
                    End If

                    Return num
                End Get
            End Property

            Public ReadOnly Property Maximum As Double
                Get
                    If WordWrapOn Then
                        Return _textView.TextSnapshot.Length
                    End If

                    Return _textView.TextSnapshot.LineCount - 1
                End Get
            End Property

            Private ReadOnly Property WordWrapOn As Boolean
                Get
                    Return (_textView.WordWrapStyle And WordWrapStyles.WordWrap) = WordWrapStyles.WordWrap
                End Get
            End Property

            Public Event MappingChanged As EventHandler

            Public Function GetBounds(span1 As Span) As DoubleSpan
                If span1.End > _textView.TextSnapshot.Length Then
                    Throw New ArgumentOutOfRangeException("span")
                End If

                Dim num As Double
                Dim num2 As Double

                If WordWrapOn Then
                    num = span1.Start
                    num2 = span1.End
                Else
                    num = _textView.TextSnapshot.GetLineFromPosition(span1.Start).LineNumber
                    num2 = _textView.TextSnapshot.GetLineFromPosition(span1.End).LineNumber
                End If

                Return New DoubleSpan(num, num2 - num)
            End Function

            Public Function GetCharacterAtCoordinate(coordinate As Double) As Integer?
                If Double.IsNaN(coordinate) Then
                    Throw New ArgumentOutOfRangeException("coordinate")
                End If

                Dim num As Integer = CInt(CLng(Fix(coordinate)) Mod Integer.MaxValue)
                If WordWrapOn Then
                    If 0 > num OrElse num > _textView.TextSnapshot.Length Then
                        Return Nothing
                    End If
                    Return num
                End If

                If 0 > num OrElse num >= _textView.TextSnapshot.LineCount Then
                    Return Nothing
                End If

                Return _textView.TextSnapshot.GetLineFromLineNumber(num).Start
            End Function

            Friend Sub New(textView As ITextView)
                _textView = textView
                AddHandler _textView.ViewportHeightChanged, AddressOf OnViewportChanged
                AddHandler _textView.ViewportWidthChanged, AddressOf OnViewportChanged
                AddHandler _textView.LayoutChanged, AddressOf OnViewportChanged
            End Sub

            Private Sub OnViewportChanged(sender As Object, e As EventArgs)
                RaiseEvent MappingChanged(sender, e)
            End Sub
        End Class

        Public Structure DoubleSpan

            Public ReadOnly Property Minimum As Double

            Public ReadOnly Property Distance As Double

            Public ReadOnly Property Maximum As Double
                Get
                    Return _Minimum + _Distance
                End Get
            End Property

            Public Sub New(min As Double, distance As Double)
                If Double.IsNaN(min) Then
                    Throw New ArgumentOutOfRangeException("minimum")
                End If

                If Double.IsNaN(distance) OrElse distance < 0.0 Then
                    Throw New ArgumentOutOfRangeException("distance")
                End If

                _Minimum = min
                _Distance = distance
            End Sub

            Public Overrides Function ToString() As String
                Return String.Format(CultureInfo.InvariantCulture, "[{0},{1}]", New Object(1) {Minimum, Maximum})
            End Function

            Public Overrides Function GetHashCode() As Integer
                Return _Minimum.GetHashCode() Xor _Distance.GetHashCode()
            End Function

            Public Overrides Function Equals(obj As Object) As Boolean
                If obj Is Nothing OrElse TypeOf obj IsNot DoubleSpan Then
                    Return False
                End If

                Return CType(obj, DoubleSpan) = Me
            End Function

            Public Shared Operator =(bounds1 As DoubleSpan, bounds2 As DoubleSpan) As Boolean
                If bounds1.Minimum = bounds2.Minimum Then
                    Return bounds1.Distance = bounds2.Distance
                End If
                Return False
            End Operator

            Public Shared Operator <>(bounds1 As DoubleSpan, bounds2 As DoubleSpan) As Boolean
                If bounds1.Minimum = bounds2.Minimum Then
                    Return bounds1.Distance <> bounds2.Distance
                End If
                Return True
            End Operator
        End Structure

        Public Const Name As String = "Avalon Vertical Scrollbar"
        Private _scrollbar As ScrollBar
        Private _textView As IAvalonTextView
        Private _scrollMap As DefaultScrollMap

        Public ReadOnly Property VisualElement As FrameworkElement Implements IAvalonTextViewMargin.VisualElement
            Get
                Return _scrollbar
            End Get
        End Property

        Public ReadOnly Property MarginSize As Double Implements ITextViewMargin.MarginSize
            Get
                Return _scrollbar.ActualWidth
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
            _scrollMap = New DefaultScrollMap(_textView)
            _scrollbar = New ScrollBar
            _scrollbar.Orientation = Orientation.Vertical
            _scrollbar.Minimum = 0.0
            OnEditorLayoutChanged(Nothing, Nothing)
            RegisterEvents()
        End Sub

        Private Sub RegisterEvents()
            AddHandler _textView.LayoutChanged, AddressOf OnEditorLayoutChanged
            AddHandler _scrollbar.Scroll, AddressOf OnVerticalScrollBarScrolled
        End Sub

        Private Sub UnRegisterEvents()
            RemoveHandler _textView.LayoutChanged, AddressOf OnEditorLayoutChanged
            RemoveHandler _scrollbar.Scroll, AddressOf OnVerticalScrollBarScrolled
        End Sub

        Private Sub OnVerticalScrollBarScrolled(sender As Object, e As ScrollEventArgs)
            _textView.OnScrollChanged(e)
            If e.Handled Then Return

            Select Case e.ScrollEventType

                Case ScrollEventType.LargeIncrement
                    _textView.ViewScroller.ScrollViewportVerticallyByPage(ScrollDirection.Down)

                Case ScrollEventType.LargeDecrement
                    _textView.ViewScroller.ScrollViewportVerticallyByPage(ScrollDirection.Up)

                Case ScrollEventType.SmallIncrement
                    _textView.ViewScroller.ScrollViewportVerticallyByLine(ScrollDirection.Down)

                Case ScrollEventType.SmallDecrement
                    _textView.ViewScroller.ScrollViewportVerticallyByLine(ScrollDirection.Up)

                Case Else
                    Dim index = _scrollMap.GetCharacterAtCoordinate(e.NewValue)
                    If index.HasValue Then
                        _textView.DisplayTextLineContainingCharacter(index.Value, 0.0, ViewRelativePosition.Top)
                    End If

            End Select

        End Sub

        Private Sub OnEditorLayoutChanged(sender As Object, e As EventArgs)
            _scrollbar.Maximum = _scrollMap.Maximum
            _scrollbar.ViewportSize = _scrollMap.ViewportSize

            Dim s = New Span(_textView.FirstVisibleCharacterPosition, 0)
            _scrollbar.Value = _scrollMap.GetBounds(s).Minimum
        End Sub
    End Class
End Namespace
