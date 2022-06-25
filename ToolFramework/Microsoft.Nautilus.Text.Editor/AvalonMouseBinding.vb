Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Input
Imports Microsoft.Nautilus.Text.Operations

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class AvalonMouseBinding

        Private _avalonTextViewHost As IAvalonTextViewHost
        Private _textView As IAvalonTextView
        Private _operations As IEditorOperations
        Private _mouseDownInsideCanvas As Boolean
        Private _mouseBindings As List(Of IMouseBinding)

        Public Sub New(avalonTextViewHost As IAvalonTextViewHost, editorOperationsProvider As IEditorOperationsProvider, mouseBindingFactories As IEnumerable(Of MouseBindingFactory))
            _avalonTextViewHost = avalonTextViewHost
            _textView = avalonTextViewHost.TextView
            _operations = editorOperationsProvider.GetEditorOperations(_textView)
            _mouseBindings = New List(Of IMouseBinding)

            For Each factory In mouseBindingFactories
                Dim associatedBinding As IMouseBinding = factory.GetAssociatedBinding(_avalonTextViewHost)
                If associatedBinding IsNot Nothing Then
                    _mouseBindings.Add(associatedBinding)
                End If
            Next

            AddHandler _textView.VisualElement.MouseLeftButtonDown, AddressOf VisualElement_MouseLeftButtonDown
            AddHandler _textView.VisualElement.MouseLeftButtonUp, AddressOf VisualElement_MouseLeftButtonUp
            AddHandler _textView.VisualElement.MouseMove, AddressOf VisualElement_MouseMove
            AddHandler _textView.VisualElement.MouseWheel, AddressOf VisualElement_MouseWheel
            AddHandler _textView.VisualElement.MouseRightButtonDown, AddressOf VisualElement_MouseRightButtonDown
            AddHandler _textView.VisualElement.MouseRightButtonUp, AddressOf VisualElement_MouseRightButtonUp
            AddHandler _textView.VisualElement.MouseEnter, AddressOf VisualElement_MouseEnter
            AddHandler _textView.VisualElement.MouseLeave, AddressOf VisualElement_MouseLeave
        End Sub

        Private Sub VisualElement_MouseWheel(sender As Object, e As MouseWheelEventArgs)
            For Each mouseBinding As IMouseBinding In _mouseBindings
                mouseBinding.PreprocessMouseWheel(e)
                If e.Handled Then
                    Exit For
                End If
            Next

            If Not e.Handled Then
                If CDbl(Math.Abs(e.Delta)) > _textView.ViewportHeight Then
                    _textView.ViewScroller.ScrollViewportVerticallyByPage(If((e.Delta < 0), ScrollDirection.Down, ScrollDirection.Up))
                Else
                    _textView.ViewScroller.ScrollViewportVerticallyByPixels(e.Delta)
                End If
            End If

            For Each mouseBinding2 As IMouseBinding In _mouseBindings
                mouseBinding2.PostprocessMouseWheel(e)
            Next
        End Sub

        Private Sub VisualElement_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
            For Each mouseBinding As IMouseBinding In _mouseBindings
                mouseBinding.PreprocessMouseLeftButtonDown(e)
                If e.Handled Then
                    Exit For
                End If
            Next

            If Not e.Handled Then
                Dim flag As Boolean = (Keyboard.Modifiers And ModifierKeys.Shift) = ModifierKeys.Shift
                _textView.VisualElement.CaptureMouse()
                Dim position As Point = e.GetPosition(_textView.VisualElement)
                position.X += _textView.ViewportLeft
                Dim textLineContainingYCoordinate As ITextLine = _textView.FormattedTextLines.GetTextLineContainingYCoordinate(position.Y)
                If textLineContainingYCoordinate IsNot Nothing Then
                    If flag Then
                        _operations.MoveCaretAndExtendSelection(textLineContainingYCoordinate, position.X)
                    Else
                        textLineContainingYCoordinate.MoveCaretToLocation(position.X)
                        _textView.Selection.Clear()
                    End If
                    _textView.Caret.EnsureVisible()
                    _textView.Caret.CapturePreferredBounds()
                    _mouseDownInsideCanvas = True
                End If
            End If

            For Each mouseBinding2 As IMouseBinding In _mouseBindings
                mouseBinding2.PostprocessMouseLeftButtonDown(e)
            Next
        End Sub

        Private Sub VisualElement_MouseMove(sender As Object, e As MouseEventArgs)
            For Each mouseBinding As IMouseBinding In _mouseBindings
                mouseBinding.PreprocessMouseMove(e)
                If e.Handled Then
                    Exit For
                End If
            Next

            If Not e.Handled AndAlso _mouseDownInsideCanvas AndAlso e.LeftButton = MouseButtonState.Pressed Then
                Dim formattedTextLines1 As IFormattedTextLineCollection = _textView.FormattedTextLines
                Dim position As Point = e.GetPosition(_textView.VisualElement)
                position.X += _textView.ViewportLeft
                Dim textLine As ITextLine
                Dim index As Integer

                If formattedTextLines1.FindTextLineIndexContainingYCoordinate(position.Y, index) Then
                    textLine = formattedTextLines1(index)
                ElseIf index = -1 Then
                    textLine = formattedTextLines1(0)
                    _textView.ViewScroller.ScrollViewportVerticallyByLine(ScrollDirection.Up)
                    index = formattedTextLines1.GetIndexOfTextLine(textLine)
                    If index > 0 Then
                        textLine = formattedTextLines1(index - 1)
                    End If
                Else
                    textLine = formattedTextLines1(formattedTextLines1.Count - 1)
                    _textView.ViewScroller.ScrollViewportVerticallyByLine(ScrollDirection.Down)
                    index = formattedTextLines1.GetIndexOfTextLine(textLine)
                    If index < formattedTextLines1.Count - 1 Then
                        textLine = formattedTextLines1(index + 1)
                    End If
                End If

                If textLine IsNot Nothing Then
                    _operations.MoveCaretAndExtendSelection(textLine, position.X)
                    _textView.Caret.EnsureVisible()
                    _textView.Caret.CapturePreferredBounds()
                End If
            End If

            For Each mouseBinding2 As IMouseBinding In _mouseBindings
                mouseBinding2.PostprocessMouseMove(e)
            Next
        End Sub

        Private Sub VisualElement_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
            For Each mouseBinding As IMouseBinding In _mouseBindings
                mouseBinding.PreprocessMouseLeftButtonUp(e)
                If e.Handled Then
                    Exit For
                End If
            Next

            If Not e.Handled Then
                _textView.VisualElement.ReleaseMouseCapture()
                Dim position As Point = e.GetPosition(_textView.VisualElement)
                position.X += _textView.ViewportLeft
                Dim textLineContainingYCoordinate As ITextLine = _textView.FormattedTextLines.GetTextLineContainingYCoordinate(position.Y)
                If textLineContainingYCoordinate IsNot Nothing AndAlso _mouseDownInsideCanvas Then
                    textLineContainingYCoordinate.MoveCaretToLocation(position.X)
                    _textView.Caret.EnsureVisible()
                    _textView.Caret.CapturePreferredBounds()
                End If
                _mouseDownInsideCanvas = False
            End If

            For Each mouseBinding2 As IMouseBinding In _mouseBindings
                mouseBinding2.PostprocessMouseLeftButtonUp(e)
            Next
        End Sub

        Private Sub VisualElement_MouseLeave(sender As Object, e As MouseEventArgs)
            For Each mouseBinding As IMouseBinding In _mouseBindings
                mouseBinding.PreprocessMouseLeave(e)
                If e.Handled Then
                    Exit For
                End If
            Next

            For Each mouseBinding2 As IMouseBinding In _mouseBindings
                mouseBinding2.PostprocessMouseLeave(e)
            Next
        End Sub

        Private Sub VisualElement_MouseEnter(sender As Object, e As MouseEventArgs)
            For Each mouseBinding As IMouseBinding In _mouseBindings
                mouseBinding.PreprocessMouseEnter(e)
                If e.Handled Then
                    Exit For
                End If
            Next

            For Each mouseBinding2 As IMouseBinding In _mouseBindings
                mouseBinding2.PostprocessMouseEnter(e)
            Next
        End Sub

        Private Sub VisualElement_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs)
            For Each mouseBinding As IMouseBinding In _mouseBindings
                mouseBinding.PreprocessMouseRightButtonUp(e)
                If e.Handled Then
                    Exit For
                End If
            Next

            For Each mouseBinding2 As IMouseBinding In _mouseBindings
                mouseBinding2.PostprocessMouseRightButtonUp(e)
            Next
        End Sub

        Private Sub VisualElement_MouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)
            For Each mouseBinding As IMouseBinding In _mouseBindings
                mouseBinding.PreprocessMouseRightButtonDown(e)
                If e.Handled Then
                    Exit For
                End If
            Next

            For Each mouseBinding2 As IMouseBinding In _mouseBindings
                mouseBinding2.PostprocessMouseRightButtonDown(e)
            Next
        End Sub

    End Class
End Namespace
