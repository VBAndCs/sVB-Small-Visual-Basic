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
    Partial Public Class ErrorListControl
        Inherits ListView
        Implements IComponentConnector

        Private _document As TextDocument

        Public Sub New(document As TextDocument)
            _document = document
            Me.InitializeComponent()
            AddHandler ItemContainerGenerator.ItemsChanged, AddressOf ErrorListControl_OnItemsChanged
        End Sub

        Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
            If e.Key = Key.Return Then
                MoveToCurrentError()
                e.Handled = True
            End If

            MyBase.OnKeyDown(e)
        End Sub

        Protected Overrides Sub OnMouseDoubleClick(e As MouseButtonEventArgs)
            MoveToCurrentError()
            MyBase.OnMouseDoubleClick(e)
        End Sub

        Public Sub SelectError(index As Integer)
            SelectedIndex = -1
            If Items.Count > 0 Then
                SelectedIndex = index
                MoveToCurrentError()
            End If
        End Sub

        Private Sub MoveToCurrentError()
            If SelectedItem Is Nothing OrElse _document.EditorControl Is Nothing Then
                Return
            End If

            Dim token = _document.ErrorTokens(SelectedIndex)
            Dim lineStart = _document.EditorControl.TextView.TextSnapshot.GetLineFromLineNumber(token.Line).Start
            _document.EditorControl.EditorOperations.Select(lineStart + token.Column, token.EndColumn - token.Column)
            _document.Focus()
        End Sub

        Private Sub ErrorListControl_OnItemsChanged(sender As Object, e As ItemsChangedEventArgs)
            If e.ItemCount = 0 Then
                DoubleAnimateHeight(0.0)
            Else
                DoubleAnimateHeight(120.0)
            End If
        End Sub

        Private Sub DoubleAnimateHeight(height As Double)
            Dim animation As New DoubleAnimation(height, New Duration(TimeSpan.FromMilliseconds(200.0)))
            BeginAnimation(HeightProperty, animation)
        End Sub

        Private Sub OnCloseClick(sender As Object, e As RoutedEventArgs)
            DoubleAnimateHeight(0.0)
        End Sub

        Public Sub Close()
            DoubleAnimateHeight(0.0)
        End Sub

    End Class
End Namespace
