Imports System.Globalization

Friend Class TypographicFeatureListItem
    Inherits TextBlock
    Implements IComparable

    Private ReadOnly _displayName As String
    Private ReadOnly _chooserProperty As DependencyProperty

    Public Sub New(ByVal displayName As String, ByVal chooserProperty As DependencyProperty)
        _displayName = displayName
        _chooserProperty = chooserProperty
        Me.Text = displayName
    End Sub

    Public ReadOnly Property ChooserProperty() As DependencyProperty
        Get
            Return _chooserProperty
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return _displayName
    End Function

    Private Function IComparable_CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo
        Return String.Compare(_displayName, obj.ToString(), True, CultureInfo.CurrentUICulture)
    End Function
End Class

