
Namespace Microsoft.SmallBasic.LanguageService
    Public Class CircularListItem
        Inherits ContentControl

        Public Shared ScaleProperty As DependencyProperty = DependencyProperty.Register("Scale", GetType(Double), GetType(CircularListItem), New PropertyMetadata(AddressOf OnScaleChanged))
        Private scaleTransform As ScaleTransform = New ScaleTransform(1.0, 1.0)
        Private parent As CircularList
        Private index As Integer

        Public Property Scale As Double
            Get
                Return GetValue(ScaleProperty)
            End Get
            Set(ByVal value As Double)
                SetValue(ScaleProperty, value)
            End Set
        End Property

        Public Sub New(ByVal parent As CircularList, ByVal index As Integer)
            LayoutTransform = scaleTransform
            Height = 16.0
            Me.parent = parent
            Me.index = index
        End Sub

        Protected Overrides Sub OnMouseDown(ByVal e As MouseButtonEventArgs)
            parent.SelectedIndex = index
            e.Handled = True
            MyBase.OnMouseDown(e)
        End Sub

        Private Shared Sub OnScaleChanged(ByVal obj As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim circularListItem = TryCast(obj, CircularListItem)
            Dim value As Double = e.NewValue
            circularListItem.scaleTransform.ScaleX = value
            circularListItem.scaleTransform.ScaleY = value
        End Sub

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class
End Namespace
