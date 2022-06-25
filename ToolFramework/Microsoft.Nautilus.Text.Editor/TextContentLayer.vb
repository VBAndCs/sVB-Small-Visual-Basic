Imports System.Windows
Imports System.Windows.Media

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class TextContentLayer
        Inherits FrameworkElement

        Public ReadOnly Property Children As VisualCollection

        Protected Overrides ReadOnly Property VisualChildrenCount As Integer
            Get
                Return _children.Count
            End Get
        End Property

        Public Sub New()
            _children = New VisualCollection(Me)
        End Sub

        Protected Overrides Function GetVisualChild(index As Integer) As Visual
            If index < 0 OrElse index > _children.Count Then
                Throw New ArgumentOutOfRangeException("index")
            End If

            Return _children(index)
        End Function

        Protected Overrides Function MeasureOverride(availableSize As Size) As Size
            Return MyBase.MeasureOverride(availableSize)
        End Function

        Protected Overrides Function ArrangeOverride(finalSize As Size) As Size
            Return MyBase.ArrangeOverride(finalSize)
        End Function

    End Class
End Namespace
