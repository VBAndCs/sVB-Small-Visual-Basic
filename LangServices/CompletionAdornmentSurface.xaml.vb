Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Threading
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.SmallBasic.Completion
Imports Microsoft.SmallBasic.Library

Namespace Microsoft.SmallBasic.LanguageService
    Partial Public Class CompletionAdornmentSurface
        Inherits Canvas
        Implements IAdornmentSurface

        Private textView As IAvalonTextView
        Private _adornment As CompletionAdornment
        Private filteredCompletionItems As List(Of CompletionItemWrapper)

        Public ReadOnly Property CanNegotiateSurfaceSpace As Boolean Implements IAdornmentSurface.CanNegotiateSurfaceSpace
            Get
                Return False
            End Get
        End Property

        Public ReadOnly Property SurfacePosition As SurfacePosition Implements IAdornmentSurface.SurfacePosition
            Get
                Return SurfacePosition.AboveText
            End Get
        End Property

        Public ReadOnly Property SurfacePanel As Panel Implements IAdornmentSurface.SurfacePanel
            Get
                Return Me
            End Get
        End Property

        Public ReadOnly Property IsAdornmentVisible As Boolean
            Get
                Return CompletionPopup.IsOpen
            End Get
        End Property

        Public ReadOnly Property Adornment As CompletionAdornment
            Get
                Return _adornment
            End Get
        End Property

        Public ReadOnly Property TextFlowDirection As FlowDirection
            Get
                Dim currentUICulture = CultureInfo.CurrentUICulture

                If Not currentUICulture.TextInfo.IsRightToLeft Then
                    Return FlowDirection.LeftToRight
                End If

                Return FlowDirection.RightToLeft
            End Get
        End Property

        Public Sub New(textView As IAvalonTextView)
            Me.textView = textView
            AddHandler Me.textView.TextBuffer.Changed, AddressOf OnTextChanged
            InitializeComponent()
            AddHandler Me.CompletionPopup.MouseDown, Sub(sender As Object, e As MouseButtonEventArgs) e.Handled = True
        End Sub

        Public Sub AddAdornment(adornment As IAdornment) Implements IAdornmentSurface.AddAdornment
            _adornment = TryCast(adornment, CompletionAdornment)
            Dispatcher.BeginInvoke(
                    DispatcherPriority.Normal,
                    CType(Function()
                              DisplayListBox()
                              Return Nothing
                          End Function, DispatcherOperationCallback),
                    Nothing
                )
        End Sub

        Public Sub RemoveAdornment(adornment As IAdornment) Implements IAdornmentSurface.RemoveAdornment
            _adornment = Nothing
            CompletionPopup.IsOpen = False
        End Sub

        Public Function GetSpaceNegotiations(textLine As ITextLine) As IList(Of SpaceNegotiation) Implements IAdornmentSurface.GetSpaceNegotiations
            Return Nothing
        End Function

        Public Function GetAdornmentsGeometry() As Geometry Implements IAdornmentSurface.GetAdornmentsGeometry
            Return Nothing
        End Function

        Public Sub FadeCompletionList()
            Dim animation As DoubleAnimation = New DoubleAnimation(0.3, New Duration(TimeSpan.FromMilliseconds(300.0)))
            popupContent.BeginAnimation(OpacityProperty, animation)
        End Sub

        Public Sub UnfadeCompletionList()
            Dim animation As DoubleAnimation = New DoubleAnimation(1.0, New Duration(TimeSpan.FromMilliseconds(300.0)))
            popupContent.BeginAnimation(OpacityProperty, animation)
        End Sub

        Private Sub OnCompletionClosed(sender As Object, e As EventArgs)
            If _adornment IsNot Nothing Then
                _adornment.Dismiss(force:=False)
            End If
        End Sub

        Private Sub CompletionPopup_MouseWheel(sender As Object, e As MouseWheelEventArgs)
            If CompletionListBox IsNot Nothing Then
                If e.Delta < 0 Then
                    CompletionListBox.MoveDown()
                Else
                    CompletionListBox.MoveUp()
                End If

                e.Handled = True
            End If
        End Sub

        Private Sub DisplayListBox()
            If _adornment IsNot Nothing Then
                Dim span As Span = _adornment.ReplaceSpan.GetSpan(textView.TextSnapshot)
                Dim textLine = textView.FormattedTextLines.GetTextLineContainingPosition(span.Start)

                If textLine IsNot Nothing Then
                    Dim characterBounds = textLine.GetCharacterBounds(span.Start)
                    Me.CompletionPopup.VerticalOffset = characterBounds.Bottom
                    CompletionPopup.HorizontalOffset = characterBounds.Left
                    UnfadeCompletionList()
                    BuildFilteredCompletionList()
                    Dim index = GetSelectedItemIndex()
                    If index > -1 Then
                        CompletionListBox.ItemsSource = filteredCompletionItems
                        CompletionPopup.IsOpen = True
                        CompletionListBox.SelectedIndex = index
                    Else
                        _adornment.Dismiss(True)
                    End If
                End If
            End If
        End Sub

        Private Sub OnTextChanged(ByVal sender As Object, ByVal e As Nautilus.Text.TextChangedEventArgs)
            If IsAdornmentVisible Then
                Dim selectedItemIndex As Integer = GetSelectedItemIndex()
                CompletionListBox.SelectedIndex = selectedItemIndex
                UnfadeCompletionList()
            End If
        End Sub

        Private Sub OnCurrentCompletionItemChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
            If CompletionListBox.SelectedItem IsNot Nothing Then
                UpdateCurrentCompletionItem(TryCast(CompletionListBox.SelectedItem, CompletionItemWrapper))
            End If
        End Sub

        Private Sub BuildFilteredCompletionList()
            filteredCompletionItems = New List(Of CompletionItemWrapper)()
            Dim completionItems = _adornment.CompletionBag.CompletionItems

            For Each item In completionItems
                If CanAddItem(item) Then
                    filteredCompletionItems.Add(New CompletionItemWrapper(item))
                End If
            Next

            Do
                Dim n = filteredCompletionItems.Count
                If n > 7 Then Exit Do

                For Each item In completionItems
                    If CanAddItem(item) Then filteredCompletionItems.Add(
                            New CompletionItemWrapper(item))
                Next
                If filteredCompletionItems.Count = n Then Exit Do
            Loop
        End Sub

        Private Function GetUniqueItems(ByVal completionItems As CompletionItem()) As List(Of CompletionItem)
            Dim dictionary As Dictionary(Of String, CompletionItem) = New Dictionary(Of String, CompletionItem)()

            For Each completionItem In completionItems
                If Not dictionary.ContainsKey(completionItem.Name) Then
                    dictionary(completionItem.Name) = completionItem
                End If
            Next

            Dim list As List(Of CompletionItem) = New List(Of CompletionItem)(dictionary.Values)
            list.Sort(New Comparison(Of CompletionItem)(AddressOf CompletionItemComparer))
            Return list
        End Function

        Private Function CanAddItem(ByVal item As CompletionItem) As Boolean
            If item.DisplayName.StartsWith("_") Then
                Return False
            End If

            If Equals(item.DisplayName, "GetHashCode") OrElse Equals(item.DisplayName, "ToString") OrElse Equals(item.DisplayName, "Equals") OrElse Equals(item.DisplayName, "GetType") Then
                Return False
            End If

            If item.MemberInfo IsNot Nothing AndAlso Equals(item.MemberInfo.Name, GetType(Primitive).Name) Then
                Return False
            End If

            Return True
        End Function

        Private Function CompletionItemComparer(ByVal item1 As CompletionItem, ByVal item2 As CompletionItem) As Integer
            Return item1.DisplayName.CompareTo(item2.DisplayName)
        End Function

        Private Function ItemWrapperComparer(ByVal item1 As CompletionItemWrapper, ByVal item2 As CompletionItemWrapper) As Integer
            Return item1.Display.CompareTo(item2.Display)
        End Function

        Private Function GetSelectedItemIndex() As Integer
            Dim text = _adornment.ReplaceSpan.GetText(_adornment.ReplaceSpan.TextBuffer.CurrentSnapshot)

            If text.Length = 0 Then Return 0

            Dim lenght = 0
            Dim result = -1

            For i = 0 To filteredCompletionItems.Count - 1
                Dim maxMatchLength As Integer = GetMaxMatchLength(filteredCompletionItems(i).Display.ToLowerInvariant(), text.ToLowerInvariant())

                If maxMatchLength > lenght Then
                    lenght = maxMatchLength
                    result = i
                End If
            Next

            If lenght < text.Length Then
                result = -1
            End If

            Return result
        End Function

        Private Function GetMaxMatchLength(ByVal s1 As String, ByVal s2 As String) As Integer
            Dim i As Integer
            i = 0

            While i < System.Math.Min(s1.Length, s2.Length) AndAlso s1(i) = s2(i)
                i += 1
            End While

            Return i
        End Function

        Private Sub OnCompletionListDoubleClicked(ByVal sender As Object, ByVal e As MouseButtonEventArgs)
            If _adornment IsNot Nothing Then
                Dim completionItemWrapper As CompletionItemWrapper = TryCast(Me.CompletionListBox.SelectedItem, CompletionItemWrapper)

                If completionItemWrapper IsNot Nothing Then
                    _adornment.AdornmentProvider.CommitItem(completionItemWrapper.CompletionItem)
                End If
            End If
        End Sub
    End Class
End Namespace
