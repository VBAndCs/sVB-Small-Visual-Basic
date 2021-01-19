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
    Public Partial Class CompletionAdornmentSurface
        Inherits Canvas
        Implements IAdornmentSurface

        Private textView As IAvalonTextView
        Private adornmentField As CompletionAdornment
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
                Return Me.completionPopup.IsOpen
            End Get
        End Property

        Public ReadOnly Property Adornment As CompletionAdornment
            Get
                Return adornmentField
            End Get
        End Property

        Public ReadOnly Property CompletionListBox As CircularList
            Get
                Return Me.completionListBoxField
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

        Public Sub New(ByVal textView As IAvalonTextView)
            Me.textView = textView
            AddHandler Me.textView.TextBuffer.Changed, AddressOf OnTextChanged
            Me.InitializeComponent()
            AddHandler Me.completionPopup.MouseDown, Sub(sender As Object, e As MouseButtonEventArgs) e.Handled = True
        End Sub

        Public Sub AddAdornment(ByVal adornment As IAdornment) Implements IAdornmentSurface.AddAdornment
            adornmentField = TryCast(adornment, CompletionAdornment)
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, CType(Function()
                                                                        DisplayListBox()
                                                                        Return Nothing
                                                                    End Function, DispatcherOperationCallback), Nothing)
        End Sub

        Public Sub RemoveAdornment(ByVal adornment As IAdornment) Implements IAdornmentSurface.RemoveAdornment
            adornmentField = Nothing
            Me.completionPopup.IsOpen = False
        End Sub

        Public Function GetSpaceNegotiations(ByVal textLine As ITextLine) As IList(Of SpaceNegotiation) Implements IAdornmentSurface.GetSpaceNegotiations
            Return Nothing
        End Function

        Public Function GetAdornmentsGeometry() As Geometry Implements IAdornmentSurface.GetAdornmentsGeometry
            Return Nothing
        End Function

        Public Sub FadeCompletionList()
            Dim animation As DoubleAnimation = New DoubleAnimation(0.3, New Duration(TimeSpan.FromMilliseconds(300.0)))
            Me.popupContent.BeginAnimation(OpacityProperty, animation)
        End Sub

        Public Sub UnfadeCompletionList()
            Dim animation As DoubleAnimation = New DoubleAnimation(1.0, New Duration(TimeSpan.FromMilliseconds(300.0)))
            Me.popupContent.BeginAnimation(OpacityProperty, animation)
        End Sub

        Private Sub OnCompletionClosed(ByVal sender As Object, ByVal e As EventArgs)
            If adornmentField IsNot Nothing Then
                adornmentField.Dismiss(force:=False)
            End If
        End Sub

        Private Sub OnMouseWheel(ByVal sender As Object, ByVal e As MouseWheelEventArgs)
            If Me.completionListBoxField IsNot Nothing Then
                If e.Delta < 0 Then
                    Me.completionListBoxField.MoveDown()
                Else
                    Me.completionListBoxField.MoveUp()
                End If

                e.Handled = True
            End If
        End Sub

        Private Sub DisplayListBox()
            If adornmentField IsNot Nothing Then
                Dim span As Span = adornmentField.ReplaceSpan.GetSpan(textView.TextSnapshot)
                Dim textLineContainingPosition = textView.FormattedTextLines.GetTextLineContainingPosition(span.Start)

                If textLineContainingPosition IsNot Nothing Then
                    Dim characterBounds = textLineContainingPosition.GetCharacterBounds(span.Start)
                    Me.completionPopup.VerticalOffset = characterBounds.Bottom
                    Me.completionPopup.HorizontalOffset = characterBounds.Left
                    UnfadeCompletionList()
                    BuildFilteredCompletionList()
                    Dim selectedItemIndex As Integer = GetSelectedItemIndex()
                    Me.completionListBoxField.ItemsSource = filteredCompletionItems
                    Me.completionPopup.IsOpen = True
                    Me.completionListBoxField.SelectedIndex = selectedItemIndex
                End If
            End If
        End Sub

        Private Sub OnTextChanged(ByVal sender As Object, ByVal e As Nautilus.Text.TextChangedEventArgs)
            If IsAdornmentVisible Then
                Dim selectedItemIndex As Integer = GetSelectedItemIndex()
                Me.completionListBoxField.SelectedIndex = selectedItemIndex
                UnfadeCompletionList()
            End If
        End Sub

        Private Sub OnCurrentCompletionItemChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
            If Me.completionListBoxField.SelectedItem IsNot Nothing Then
                CompilerService.UpdateCurrentCompletionItem(TryCast(Me.completionListBoxField.SelectedItem, CompletionItemWrapper))
            End If
        End Sub

        Private Sub BuildFilteredCompletionList()
            filteredCompletionItems = New List(Of CompletionItemWrapper)()
            Dim completionItems = adornmentField.CompletionBag.CompletionItems

            For Each item In completionItems

                If CanAddItem(item) Then
                    filteredCompletionItems.Add(New CompletionItemWrapper(item))
                End If
            Next

            While filteredCompletionItems.Count < 8 AndAlso completionItems.Count > 0

                For Each item2 In completionItems

                    If CanAddItem(item2) Then
                        filteredCompletionItems.Add(New CompletionItemWrapper(item2))
                    End If
                Next
            End While
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
            Dim text = adornmentField.ReplaceSpan.GetText(adornmentField.ReplaceSpan.TextBuffer.CurrentSnapshot)

            If text.Length = 0 Then
                Return 0
            End If

            Dim num = 0
            Dim result = -1

            For i = 0 To filteredCompletionItems.Count - 1
                Dim maxMatchLength As Integer = GetMaxMatchLength(filteredCompletionItems(i).Display.ToLowerInvariant(), text.ToLowerInvariant())

                If maxMatchLength > num Then
                    num = maxMatchLength
                    result = i
                End If
            Next

            If num < text.Length Then
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
            If adornmentField IsNot Nothing Then
                Dim completionItemWrapper As CompletionItemWrapper = TryCast(Me.completionListBoxField.SelectedItem, CompletionItemWrapper)

                If completionItemWrapper IsNot Nothing Then
                    adornmentField.AdornmentProvider.CommitItem(completionItemWrapper.CompletionItem)
                End If
            End If
        End Sub
    End Class
End Namespace
