
Namespace Microsoft.SmallBasic.LanguageService
    Public Class CircularListItem
        Inherits ContentControl

        Public Shared ScaleProperty As DependencyProperty = DependencyProperty.Register("Scale", GetType(Double), GetType(CircularListItem), New PropertyMetadata(AddressOf OnScaleChanged))
        Private scaleTransform As ScaleTransform = New ScaleTransform(1.0, 1.0)
        Private _parent As CircularList
        Private index As Integer

        Public Property Scale As Double
            Get
                Return GetValue(ScaleProperty)
            End Get
            Set(ByVal value As Double)
                SetValue(ScaleProperty, value)
            End Set
        End Property

        Public Sub New(parent As CircularList, index As Integer)
            LayoutTransform = scaleTransform
            Height = 16.0
            Me._parent = parent
            Me.index = index
        End Sub

        Protected Overrides Sub OnMouseDown(e As MouseButtonEventArgs)
            _parent.SelectedIndex = index
            e.Handled = True
            MyBase.OnMouseDown(e)
        End Sub

        Private Shared Sub OnScaleChanged(obj As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim circularListItem = TryCast(obj, CircularListItem)
            Dim value As Double = e.NewValue
            circularListItem.scaleTransform.ScaleX = value
            circularListItem.scaleTransform.ScaleY = value
        End Sub

    End Class
End Namespace
