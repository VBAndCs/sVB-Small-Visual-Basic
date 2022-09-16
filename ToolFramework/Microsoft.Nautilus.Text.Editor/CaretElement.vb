Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Animation

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class CaretElement
        Inherits UIElement
        Implements ITextCaret

        Private NotInheritable Class NativeMethods
            <DllImport("user32.dll")>
            Friend Shared Function GetCaretBlinkTime() As Integer
            End Function
        End Class

        Private Const _bidiCaretIndicatorWidth As Double = 2.0
        Private Const _bidiIndicatorHeightRatio As Double = 10.0
        Public Const CaretHorizontalPadding As Double = 2.0
        Public Const HorizontalScrollbarPadding As Double = 10.0
        Private _caretBrush As Brush
        Private _blinkAnimationClock As AnimationClock
        Private _insertionPoint As ITextPoint
        Private _avalonTextView As AvalonTextView
        Private _capturedCharacterBounds As New TextBounds(0.0, 0.0, 0.0, 0.0)
        Private _ensureVisiblePending As Boolean
        Private _updateNeeded As Boolean = True
        Private _oldHeight As Double = -1.0
        Private _bounds As New Rect(0.0, 0.0, SystemParameters.CaretWidth, 0.0)
        Private _caretGeometry As Geometry

        Public ReadOnly Property Left As Double Implements ITextCaret.Left
            Get
                If _updateNeeded Then
                    UpdateCaret()
                End If

                Return _bounds.Left
            End Get
        End Property

        Public Property Width As Double Implements ITextCaret.Width
            Get
                Return _bounds.Width
            End Get

            Set(value As Double)
                If value < 0.0 OrElse Double.IsNaN(value) Then
                    Throw New ArgumentOutOfRangeException("value")
                End If

                _bounds.Width = value
                InvalidateVisual()
                _updateNeeded = True
                _oldHeight = -1.0
            End Set
        End Property

        Public ReadOnly Property Right As Double Implements ITextCaret.Right
            Get
                If _updateNeeded Then
                    UpdateCaret()
                End If

                Return _bounds.Right
            End Get
        End Property

        Public ReadOnly Property Top As Double Implements ITextCaret.Top
            Get
                If _updateNeeded Then
                    UpdateCaret()
                End If

                Return _bounds.Top
            End Get
        End Property

        Public ReadOnly Property Height As Double Implements ITextCaret.Height
            Get
                If _updateNeeded Then
                    UpdateCaret()
                End If

                Return _bounds.Height
            End Get
        End Property

        Public ReadOnly Property Bottom As Double Implements ITextCaret.Bottom
            Get
                If _updateNeeded Then
                    UpdateCaret()
                End If

                Return _bounds.Bottom
            End Get
        End Property

        Public ReadOnly Property Placement As CaretPlacement

        Public ReadOnly Property Position As ICaretPosition Implements ITextCaret.Position
            Get
                Dim position1 As Integer = _insertionPoint.GetPosition(_avalonTextView.TextSnapshot)
                If _Placement = CaretPlacement.TrailingEdgeOfCharacter AndAlso position1 > 0 Then
                    Return CreateCaretPosition(position1 - 1, CaretPlacement.TrailingEdgeOfCharacter)
                End If

                Return CreateCaretPosition(position1, CaretPlacement.LeadingEdgeOfCharacter)
            End Get
        End Property

        Public ReadOnly Property PreferredBounds As TextBounds Implements ITextCaret.PreferredBounds
            Get
                Return _capturedCharacterBounds
            End Get
        End Property

        Private ReadOnly Property CaretCharacterWidth As Double
            Get
                Return If(_avalonTextView.FormattedTextLines.GetTextLineContainingPosition(Position.TextInsertionIndex)?.GetCharacterBounds(Position.TextInsertionIndex).Width, 0.0)
            End Get
        End Property

        Public Event PositionChanged As EventHandler(Of CaretPositionChangedEventArgs) Implements ITextCaret.PositionChanged

        Private Function CreateCaretPosition(characterIndex As Integer, caretPlacement1 As CaretPlacement) As ICaretPosition
            Dim textElementSpan As Span = _avalonTextView.GetTextElementSpan(characterIndex)
            Return New CaretPosition(textElementSpan.Start,
                      If(caretPlacement1 = CaretPlacement.LeadingEdgeOfCharacter, textElementSpan.Start, textElementSpan.End),
                      If(textElementSpan.Length <> 0, caretPlacement1, CaretPlacement.LeadingEdgeOfCharacter))
        End Function

        Public Sub New(textView As AvalonTextView)
            _avalonTextView = textView
            _Placement = CaretPlacement.LeadingEdgeOfCharacter
            _insertionPoint = textView.TextSnapshot.CreateTextPoint(0, TrackingMode.Positive)
            Dim black1 As Color = Colors.Black
            _caretBrush = New SolidColorBrush(black1)
            If _caretBrush.CanFreeze Then _caretBrush.Freeze()

            Dim doubleAnimationUsingKeyFrames1 As New DoubleAnimationUsingKeyFrames With {
                .BeginTime = New TimeSpan(0L),
                .RepeatBehavior = RepeatBehavior.Forever,
                .KeyFrames = New DoubleKeyFrameCollection From {
                     CType(New DiscreteDoubleKeyFrame(1.0, KeyTime.FromPercent(0.0)), DoubleKeyFrame)
                }
            }

            Dim num As Integer = NativeMethods.GetCaretBlinkTime()
            If num > 0 Then
                doubleAnimationUsingKeyFrames1.KeyFrames.Add(New DiscreteDoubleKeyFrame(0.0, KeyTime.FromPercent(0.5)))
            Else
                num = 500
            End If

            doubleAnimationUsingKeyFrames1.Duration = New Duration(New TimeSpan(0, 0, 0, 0, num * 2))
            _blinkAnimationClock = doubleAnimationUsingKeyFrames1.CreateClock()
            ApplyAnimationClock(UIElement.OpacityProperty, _blinkAnimationClock)
            AddHandler MyBase.IsVisibleChanged, AddressOf OnVisibleChanged
        End Sub

        Public Sub EnsureVisible() Implements ITextCaret.EnsureVisible
            If _avalonTextView.FormattedTextLines.Count = 0 Then
                _ensureVisiblePending = True
                Return
            End If

            _ensureVisiblePending = False
            Dim span1 As New Span(Position.TextInsertionIndex, 0)
            If Not _avalonTextView.ViewScroller.EnsureSpanVisible(span1, 20.0 + Width, 20.0) Then
                _avalonTextView.ViewScroller.EnsureSpanVisible(span1, 0.0, 0.0)
            End If
        End Sub

        Public Function MoveTo(characterIndex As Integer) As ICaretPosition Implements ITextCaret.MoveTo
            Return MoveTo(characterIndex, CaretPlacement.LeadingEdgeOfCharacter)
        End Function

        Public Function MoveTo(characterIndex As Integer, caretPlacement As CaretPlacement) As ICaretPosition Implements ITextCaret.MoveTo
            Dim position1 = Position
            If characterIndex > _avalonTextView.TextSnapshot.Length Then
                characterIndex = _avalonTextView.TextSnapshot.Length
            End If
            Dim textElementSpan = _avalonTextView.GetTextElementSpan(characterIndex)
            _Placement = caretPlacement

            If _Placement = CaretPlacement.TrailingEdgeOfCharacter AndAlso characterIndex < _avalonTextView.TextSnapshot.Length Then
                _insertionPoint = _avalonTextView.TextSnapshot.CreateTextPoint(textElementSpan.End, TrackingMode.Positive)
            Else
                _insertionPoint = _avalonTextView.TextSnapshot.CreateTextPoint(textElementSpan.Start, TrackingMode.Positive)
            End If

            Dim position2 As ICaretPosition = Position
            _ensureVisiblePending = False
            InvalidateVisual()
            _updateNeeded = True

            RaiseEvent PositionChanged(Me, New CaretPositionChangedEventArgs(_avalonTextView, position1, position2))
            Return position2
        End Function

        Public Function MoveToNextCaretPosition() As ICaretPosition Implements ITextCaret.MoveToNextCaretPosition
            Dim position1 As ICaretPosition = Position
            If position1.TextInsertionIndex >= _avalonTextView.TextSnapshot.Length Then
                Return position1
            End If

            Dim [end] As Integer = _avalonTextView.TextSnapshot.GetLineFromPosition(position1.TextInsertionIndex).End
            If position1.TextInsertionIndex >= [end] Then
                Return MoveTo(_avalonTextView.TextSnapshot.GetLineFromPosition([end]).EndIncludingLineBreak, CaretPlacement.LeadingEdgeOfCharacter)
            End If

            Return MoveTo(_avalonTextView.GetTextElementSpan(position1.TextInsertionIndex).Start, CaretPlacement.TrailingEdgeOfCharacter)
        End Function

        Public Function MoveToPreviousCaretPosition() As ICaretPosition Implements ITextCaret.MoveToPreviousCaretPosition
            Dim position1 As ICaretPosition = Position
            Dim textInsertionIndex1 As Integer = position1.TextInsertionIndex
            If textInsertionIndex1 = 0 Then
                Return position1
            End If
            If textInsertionIndex1 = _avalonTextView.TextSnapshot.GetLineFromPosition(textInsertionIndex1).Start Then
                Dim [end] As Integer = _avalonTextView.TextSnapshot.GetLineFromPosition(textInsertionIndex1 - 1).End
                If textInsertionIndex1 > [end] Then
                    Return MoveTo([end], CaretPlacement.LeadingEdgeOfCharacter)
                End If
                Return MoveTo(textInsertionIndex1, CaretPlacement.TrailingEdgeOfCharacter)
            End If
            Dim textElementSpan As Span = _avalonTextView.GetTextElementSpan(If((position1.CharacterIndex > 0), (position1.CharacterIndex - 1), 0))
            Return MoveTo(If((position1.Placement <> CaretPlacement.TrailingEdgeOfCharacter), textElementSpan.Start, (If((position1.CharacterIndex = 0), textElementSpan.Start, textElementSpan.End))), CaretPlacement.LeadingEdgeOfCharacter)
        End Function

        Public Sub CapturePreferredBounds() Implements ITextCaret.CapturePreferredBounds
            _capturedCharacterBounds = New TextBounds(Left, Top, CaretCharacterWidth, Height)
        End Sub

        Public Sub CapturePreferredHorizontalBounds() Implements ITextCaret.CapturePreferredHorizontalBounds
            _capturedCharacterBounds = New TextBounds(Left, _capturedCharacterBounds.Top, CaretCharacterWidth, _capturedCharacterBounds.Height)
        End Sub

        Public Sub CapturePreferredVerticalBounds() Implements ITextCaret.CapturePreferredVerticalBounds
            _capturedCharacterBounds = New TextBounds(_capturedCharacterBounds.Left, Top, _capturedCharacterBounds.Width, Height)
        End Sub

        Protected Overrides Sub OnRender(drawingContext As DrawingContext)
            If _updateNeeded Then UpdateCaret()
            MyBase.OnRender(drawingContext)
            If _bounds.Height <> 0.0 AndAlso MyBase.Visibility = Visibility.Visible Then
                drawingContext.DrawGeometry(_caretBrush, Nothing, _caretGeometry)
            Else
                _blinkAnimationClock.Controller.Stop()
            End If
        End Sub

        Private Sub ConstructCaretGeometry(newHeight As Double)
            Dim pathGeometry1 As New PathGeometry
            pathGeometry1.AddGeometry(New RectangleGeometry(New Rect(0.0, 0.0, Width, newHeight)))
            If InputLanguageManager.Current.CurrentInputLanguage.TextInfo.IsRightToLeft Then
                Dim pathFigure1 As New PathFigure
                pathFigure1.StartPoint = New Point(0.0, 0.0)
                pathFigure1.Segments.Add(New LineSegment(New Point(-2.0, 0.0), isStroked:=True))
                pathFigure1.Segments.Add(New LineSegment(New Point(0.0, newHeight / 10.0), isStroked:=True))
                pathFigure1.IsClosed = True
                pathGeometry1.Figures.Add(pathFigure1)
            End If
            _caretGeometry = pathGeometry1
            _oldHeight = newHeight
        End Sub

        Private Sub UpdateCaret()
            _updateNeeded = False
            Dim textLineContainingPosition As ITextLine = _avalonTextView.FormattedTextLines.GetTextLineContainingPosition(Position.CharacterIndex)

            If textLineContainingPosition Is Nothing Then
                _bounds = New Rect(0.0, 0.0, Width, 0.0)
                Return
            End If

            If textLineContainingPosition.Height <> _oldHeight Then
                ConstructCaretGeometry(textLineContainingPosition.Height)
            End If

            Dim characterBounds As TextBounds = textLineContainingPosition.GetCharacterBounds(Position.CharacterIndex)
            _bounds = New Rect(If((Placement = CaretPlacement.LeadingEdgeOfCharacter), characterBounds.Left, characterBounds.Right), textLineContainingPosition.Top, Width, textLineContainingPosition.Height)
            _caretGeometry.Transform = New TranslateTransform(_bounds.Left, _bounds.Top)
            _blinkAnimationClock.Controller.Begin()
        End Sub

        Private Sub OnLayoutChanged(sender As Object, e As TextViewLayoutChangedEventArgs)
            If _ensureVisiblePending Then
                EnsureVisible()
            End If
            If e.ChangeSpan.HasValue Then
                InvalidateVisual()
                _updateNeeded = True
            End If
        End Sub

        Private Sub OnVisibleChanged(sender As Object, e As DependencyPropertyChangedEventArgs)
            If CBool(e.NewValue) Then
                AddHandler _avalonTextView.LayoutChanged, AddressOf OnLayoutChanged
                AddHandler InputLanguageManager.Current.InputLanguageChanged, AddressOf OnInputLanguageChanged
                InvalidateVisual()
                _updateNeeded = True
                _oldHeight = -1.0
            Else
                RemoveHandler _avalonTextView.LayoutChanged, AddressOf OnLayoutChanged
                RemoveHandler InputLanguageManager.Current.InputLanguageChanged, AddressOf OnInputLanguageChanged
            End If
        End Sub

        Private Sub OnInputLanguageChanged(sender As Object, e As InputLanguageEventArgs)
            InvalidateVisual()
            _updateNeeded = True
            _oldHeight = -1.0
        End Sub
    End Class
End Namespace
