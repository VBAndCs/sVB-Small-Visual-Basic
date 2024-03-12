Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports System.Windows
Imports System.Windows.Automation.Peers
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports Microsoft.Nautilus.Core
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text.Editor.Automation.Implementation
Imports Microsoft.Nautilus.Text.Operations
Imports ImportantMarginInfo = System.ComponentModel.Composition.ImportInfo(Of Microsoft.Nautilus.Text.Editor.AvalonTextViewMarginFactory, Microsoft.Nautilus.Text.Editor.IAvalonTextViewMarginFactoryMetadata)

Namespace Microsoft.Nautilus.Text.Editor
    Public NotInheritable Class AvalonTextViewHost
        Inherits ContentControl
        Implements IAvalonTextViewHost

        Private _viewTransform As ScaleTransform
        Private _avalonTextViewMarginFactories As IEnumerable(Of ImportantMarginInfo)
        Private _outerGrid As Grid
        Private _horizontalScrollBar As ITextViewMargin
        Private _verticalScrollBar As ITextViewMargin
        Private _editorMouseBinding As AvalonMouseBinding
        Private _undoHistory As UndoHistory
        Private _editorOperations As IEditorOperations
        Private _marginGrids As Grid()
        Private _extensibleMargins As Dictionary(Of String, IAvalonTextViewMargin)
        Private _keyboardFilters As IEnumerable(Of KeyboardFilter)
        Private _orderer As IOrderer

        Public Property IsHorizontalScrollBarVisible As Boolean Implements IAvalonTextViewHost.IsHorizontalScrollBarVisible
            Get
                If _horizontalScrollBar Is Nothing Then
                    Return False
                End If

                Return _horizontalScrollBar.MarginVisible
            End Get

            Set(value As Boolean)
                If _horizontalScrollBar IsNot Nothing AndAlso value <> _horizontalScrollBar.MarginVisible Then
                    _horizontalScrollBar.MarginVisible = value
                End If
            End Set
        End Property

        Public Property IsVerticalScrollBarVisible As Boolean Implements IAvalonTextViewHost.IsVerticalScrollBarVisible
            Get
                If _verticalScrollBar Is Nothing Then
                    Return False
                End If

                Return _verticalScrollBar.MarginVisible
            End Get

            Set(value As Boolean)
                If _verticalScrollBar IsNot Nothing AndAlso value <> _verticalScrollBar.MarginVisible Then
                    _verticalScrollBar.MarginVisible = value
                End If
            End Set
        End Property

        Public Property IsReadOnly As Boolean Implements IAvalonTextViewHost.IsReadOnly

        Public ReadOnly Property TextView As IAvalonTextView Implements IAvalonTextViewHost.TextView

        Public ReadOnly Property HostControl As Control Implements IAvalonTextViewHost.HostControl
            Get
                Return Me
            End Get
        End Property

        Public Property ScaleFactor As Double Implements IAvalonTextViewHost.ScaleFactor
            Get
                Return _viewTransform.ScaleX
            End Get

            Set(value As Double)
                Dim n = Math.Min(Math.Max(value, 0.1), 10.0)
                _viewTransform.ScaleY = n
                _viewTransform.ScaleX = n
            End Set
        End Property

        Public Property WordWrapStyle As WordWrapStyles Implements IAvalonTextViewHost.WordWrapStyle
            Get
                Return _TextView.WordWrapStyle
            End Get

            Set(value As WordWrapStyles)
                _textView.WordWrapStyle = value
                IsHorizontalScrollBarVisible = (_textView.WordWrapStyle And WordWrapStyles.WordWrap) <> WordWrapStyles.WordWrap
            End Set
        End Property

        Public Sub New(
                      undoHistoryRegistry As IUndoHistoryRegistry,
                      orderer As IOrderer,
                      textView As IAvalonTextView,
                      avalonTextViewMarginFactories As IEnumerable(Of ImportantMarginInfo),
                      editorOperationsProvider As IEditorOperationsProvider,
                      mouseBindingFactories As IEnumerable(Of MouseBindingFactory),
                      keyboardFilters As IEnumerable(Of KeyboardFilter))

            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            If undoHistoryRegistry Is Nothing Then
                Throw New ArgumentNullException("undoHistoryRegistry")
            End If

            _orderer = orderer
            _TextView = textView
            _editorMouseBinding = New AvalonMouseBinding(Me, editorOperationsProvider, mouseBindingFactories)
            _undoHistory = undoHistoryRegistry.RegisterHistory(textView.TextBuffer)
            _editorOperations = editorOperationsProvider.GetEditorOperations(textView)
            _avalonTextViewMarginFactories = avalonTextViewMarginFactories
            _keyboardFilters = keyboardFilters
            InitializeComponent()
            _viewTransform = New ScaleTransform(1.0, 1.0)
            _TextView.VisualElement.LayoutTransform = _viewTransform
            AddHandler _TextView.VisualElement.PreviewKeyDown, AddressOf OnPreviewKeyDown
            AddHandler _TextView.VisualElement.PreviewKeyUp, AddressOf OnPreviewKeyUp
            AddHandler _TextView.VisualElement.PreviewTextInput, AddressOf OnPreviewTextInput
            AddHandler _TextView.VisualElement.GotKeyboardFocus, AddressOf OnGotKeyboardFocus
            AddHandler _TextView.VisualElement.LostKeyboardFocus, AddressOf OnLostKeyboardFocus
            AddHandler _TextView.VisualElement.KeyDown, AddressOf OnKeyDown
            AddHandler _TextView.VisualElement.KeyUp, AddressOf OnKeyUp
            AddHandler _TextView.VisualElement.TextInput, AddressOf OnTextInput
        End Sub

        Public Function GetTextViewMargin(marginName As String) As ITextViewMargin Implements IAvalonTextViewHost.GetTextViewMargin
            If marginName Is Nothing Then
                Throw New ArgumentNullException("marginName")
            End If

            Dim value As IAvalonTextViewMargin = Nothing
            If Not _extensibleMargins.TryGetValue(marginName, value) Then
                Return Nothing
            End If
            Return value
        End Function

        Private Sub InitializeComponent()
            MyBase.Focusable = False
            _outerGrid = New Grid
            Dim columnDefinition1 As New ColumnDefinition
            columnDefinition1.Width = New GridLength(0.0, GridUnitType.Auto)
            Dim columnDefinition2 As New ColumnDefinition
            columnDefinition2.Width = New GridLength(0.0, GridUnitType.Auto)
            columnDefinition2.MinWidth = 10.0
            Dim value As New ColumnDefinition
            Dim columnDefinition3 As New ColumnDefinition
            columnDefinition3.Width = New GridLength(0.0, GridUnitType.Auto)
            _outerGrid.ColumnDefinitions.Add(columnDefinition1)
            _outerGrid.ColumnDefinitions.Add(columnDefinition2)
            _outerGrid.ColumnDefinitions.Add(value)
            _outerGrid.ColumnDefinitions.Add(columnDefinition3)
            Dim rowDefinition1 As New RowDefinition
            rowDefinition1.Height = New GridLength(0.0, GridUnitType.Auto)
            Dim value2 As New RowDefinition
            Dim rowDefinition2 As New RowDefinition
            rowDefinition2.Height = New GridLength(0.0, GridUnitType.Auto)
            _outerGrid.RowDefinitions.Add(rowDefinition1)
            _outerGrid.RowDefinitions.Add(value2)
            _outerGrid.RowDefinitions.Add(rowDefinition2)
            MyBase.Content = _outerGrid
            Grid.SetRow(_textView.VisualElement, 1)
            Grid.SetColumn(_textView.VisualElement, 2)
            _outerGrid.Children.Add(_textView.VisualElement)
            SetupMargins()
        End Sub

        Private Sub SetupMargins()
            _marginGrids = New Grid(3) {}
            _extensibleMargins = New Dictionary(Of String, IAvalonTextViewMargin)
            _marginGrids(2) = New Grid
            Dim value As New ColumnDefinition
            _marginGrids(2).ColumnDefinitions.Add(value)
            Grid.SetRow(_marginGrids(2), 0)
            Grid.SetColumnSpan(_marginGrids(2), 2)
            Grid.SetColumn(_marginGrids(2), 1)
            _outerGrid.Children.Add(_marginGrids(2))
            _marginGrids(3) = New Grid
            value = New ColumnDefinition
            _marginGrids(3).ColumnDefinitions.Add(value)
            Grid.SetRow(_marginGrids(3), 2)
            Grid.SetColumnSpan(_marginGrids(3), 2)
            Grid.SetColumn(_marginGrids(3), 1)
            _outerGrid.Children.Add(_marginGrids(3))
            _marginGrids(0) = New Grid
            Dim value2 As New RowDefinition
            _marginGrids(0).RowDefinitions.Add(value2)
            Grid.SetRow(_marginGrids(0), 1)
            Grid.SetColumn(_marginGrids(0), 0)
            _outerGrid.Children.Add(_marginGrids(0))
            _marginGrids(1) = New Grid
            value2 = New RowDefinition
            _marginGrids(1).RowDefinitions.Add(value2)
            Grid.SetRow(_marginGrids(1), 1)
            Grid.SetColumn(_marginGrids(1), 3)
            _outerGrid.Children.Add(_marginGrids(1))
            SetupExtensibleMargins()
            _horizontalScrollBar = GetTextViewMargin("Avalon Horizontal Scrollbar")
            _verticalScrollBar = GetTextViewMargin("Avalon Vertical Scrollbar")
            IsHorizontalScrollBarVisible = (_textView.WordWrapStyle And WordWrapStyles.WordWrap) <> WordWrapStyles.WordWrap
            IsVerticalScrollBarVisible = True
        End Sub

        Private Function IsMarginPlacementHorizontal(placement As MarginPlacement) As Boolean
            Return placement = MarginPlacement.Top OrElse placement = MarginPlacement.Bottom
        End Function

        Private Function IsMarginPlacementOrderReversed(placement As MarginPlacement) As Boolean
            Return placement = MarginPlacement.Top OrElse placement = MarginPlacement.Left
        End Function

        Private Sub SetupExtensibleMargin(placement As MarginPlacement)
            Dim marginInfo As New List(Of ImportantMarginInfo)
            For Each marginFactory In _avalonTextViewMarginFactories
                If marginFactory.Metadata.MarginPlacement = placement Then
                    marginInfo.Add(marginFactory)
                End If
            Next

            marginInfo = _orderer.Order(marginInfo).ToList()
            For i = 0 To marginInfo.Count - 1
                If IsMarginPlacementHorizontal(placement) Then
                    _marginGrids(placement).RowDefinitions.Add(New RowDefinition)
                Else
                    _marginGrids(placement).ColumnDefinitions.Add(New ColumnDefinition)
                End If

                Dim margin = marginInfo(i).GetBoundValue().CreateAvalonTextViewMargin(_TextView)
                Grid.SetColumn(margin.VisualElement, _marginGrids(placement).ColumnDefinitions.Count - 1)
                Grid.SetRow(margin.VisualElement, _marginGrids(placement).RowDefinitions.Count - 1)
                _marginGrids(placement).Children.Add(margin.VisualElement)
                _extensibleMargins.Add(marginInfo(i).Metadata.Name, margin)
            Next
        End Sub

        Private Sub SetupExtensibleMargins()
            SetupExtensibleMargin(MarginPlacement.Top)
            SetupExtensibleMargin(MarginPlacement.Bottom)
            SetupExtensibleMargin(MarginPlacement.Left)
            SetupExtensibleMargin(MarginPlacement.Right)
        End Sub

        Protected Overrides Function OnCreateAutomationPeer() As AutomationPeer
            Return New TextViewHostAutomationPeer(Me)
        End Function

        Private Overloads Sub OnTextInput(sender As Object, e As TextCompositionEventArgs)
            For Each keyboardFilter In _keyboardFilters
                If Not e.Handled OrElse keyboardFilter.IsApplicableToHandledEvents Then
                    keyboardFilter.TextInput(_textView, e)
                End If
            Next
        End Sub

        Private Overloads Sub OnPreviewTextInput(sender As Object, e As TextCompositionEventArgs)
            For Each keyboardFilter1 As KeyboardFilter In _keyboardFilters
                If Not e.Handled OrElse keyboardFilter1.IsApplicableToHandledEvents Then
                    keyboardFilter1.PreviewTextInput(_textView, e)
                End If
            Next
        End Sub

        Private Overloads Sub OnKeyUp(sender As Object, e As KeyEventArgs)
            For Each keyboardFilter1 As KeyboardFilter In _keyboardFilters
                If Not e.Handled OrElse keyboardFilter1.IsApplicableToHandledEvents Then
                    keyboardFilter1.KeyUp(_textView, e)
                End If
            Next
        End Sub

        Private Overloads Sub OnKeyDown(sender As Object, e As KeyEventArgs)
            For Each keyboardFilter In _keyboardFilters
                If Not e.Handled OrElse keyboardFilter.IsApplicableToHandledEvents Then
                    keyboardFilter.KeyDown(_textView, e)
                End If
            Next
        End Sub

        Private Overloads Sub OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
            For Each keyboardFilter In _keyboardFilters
                If Not e.Handled OrElse keyboardFilter.IsApplicableToHandledEvents Then
                    keyboardFilter.LostKeyboardFocus(_textView, e)
                End If
            Next
        End Sub

        Private Overloads Sub OnPreviewKeyUp(sender As Object, e As KeyEventArgs)
            For Each keyboardFilter1 As KeyboardFilter In _keyboardFilters
                If Not e.Handled OrElse keyboardFilter1.IsApplicableToHandledEvents Then
                    keyboardFilter1.PreviewKeyUp(_textView, e)
                End If
            Next
        End Sub

        Private Overloads Sub OnGotKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
            For Each keyboardFilter1 As KeyboardFilter In _keyboardFilters
                If Not e.Handled OrElse keyboardFilter1.IsApplicableToHandledEvents Then
                    keyboardFilter1.GotKeyboardFocus(_textView, e)
                End If
            Next
        End Sub

        Private Overloads Sub OnPreviewKeyDown(sender As Object, e As KeyEventArgs)
            For Each keyboardFilter1 As KeyboardFilter In _keyboardFilters
                If Not e.Handled OrElse keyboardFilter1.IsApplicableToHandledEvents Then
                    keyboardFilter1.PreviewKeyDown(_textView, e)
                End If
            Next
        End Sub

    End Class
End Namespace
