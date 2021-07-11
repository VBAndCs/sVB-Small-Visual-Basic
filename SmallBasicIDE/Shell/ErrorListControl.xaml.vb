Imports Microsoft.SmallBasic.Documents
Imports System
Imports System.Text.RegularExpressions
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Windows.Markup
Imports System.Windows.Media.Animation

Namespace Microsoft.SmallBasic.Shell
    Public Partial Class ErrorListControl
        Inherits ListView
        Implements IComponentConnector

        Private _document As TextDocument

        Public Sub New(ByVal document As TextDocument)
            _document = document
            Me.InitializeComponent()
            AddHandler ItemContainerGenerator.ItemsChanged, AddressOf OnItemsChanged
        End Sub

        Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)
            If e.Key = Key.Return Then
                MoveToCurrentError()
                e.Handled = True
            End If

            MyBase.OnKeyDown(e)
        End Sub

        Protected Overrides Sub OnMouseDoubleClick(ByVal e As MouseButtonEventArgs)
            MoveToCurrentError()
            MyBase.OnMouseDoubleClick(e)
        End Sub

        Private Sub MoveToCurrentError()
            If SelectedItem Is Nothing OrElse _document.EditorControl Is Nothing Then
                Return
            End If

            Dim text As String = TryCast(SelectedItem, String)

            If Not Equals(text, Nothing) Then
                Dim regex As Regex = New Regex("([0-9]*),([0-9]*): w*")
                Dim match = regex.Match(text)

                If match.Success Then
                    Dim line = Integer.Parse(match.Groups(1).Value) - 1
                    Dim column = Integer.Parse(match.Groups(2).Value) - 1
                    SetCaret(line, column)
                End If
            End If
        End Sub

        Private Sub SetCaret(line As Integer, column As Integer)
            If line < 0 Then Return

            Dim currentSnapshot = _document.TextBuffer.CurrentSnapshot

            If line < currentSnapshot.LineCount Then
                Dim lineFromLineNumber = currentSnapshot.GetLineFromLineNumber(line)

                If column < lineFromLineNumber.LengthIncludingLineBreak Then
                    Dim characterIndex = lineFromLineNumber.Start + column
                    _document.EditorControl.TextView.Caret.MoveTo(characterIndex)
                    _document.EditorControl.TextView.Caret.EnsureVisible()
                    _document.EditorControl.TextView.VisualElement.Focus()
                End If
            End If
        End Sub

        Private Sub OnItemsChanged(ByVal sender As Object, ByVal e As ItemsChangedEventArgs)
            If e.ItemCount = 0 Then
                DoubleAnimateHeight(0.0)
            Else
                DoubleAnimateHeight(120.0)
            End If
        End Sub

        Private Sub DoubleAnimateHeight(ByVal height As Double)
            Dim animation As DoubleAnimation = New DoubleAnimation(height, New Duration(TimeSpan.FromMilliseconds(200.0)))
            BeginAnimation(HeightProperty, animation)
        End Sub

        Private Sub OnCloseClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
            DoubleAnimateHeight(0.0)
        End Sub
    End Class
End Namespace
