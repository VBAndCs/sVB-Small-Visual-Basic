Imports Microsoft.SmallBasic.com.smallbasic
Imports System
Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Markup
Imports System.Windows.Media
Imports System.Windows.Shapes

Namespace Microsoft.SmallBasic.Utility
    Public Partial Class RatingControl
        Inherits Control
        Implements IComponentConnector

        Public Shared ReadOnly ProgramIdProperty As DependencyProperty = DependencyProperty.Register("ProgramId", GetType(String), GetType(RatingControl))
        Public Shared ReadOnly RatingProperty As DependencyProperty = DependencyProperty.Register("Rating", GetType(Double), GetType(RatingControl), New PropertyMetadata(AddressOf OnRatingChanged))
        Private filledColor As Color = Colors.Yellow
        Private unfilledColor As Color = Color.FromArgb(Byte.MaxValue, 221, 221, 221)
        Private filledBrush As Brush
        Private unfilledBrush As Brush

        Public Property ProgramId As String
            Get
                Return CStr(GetValue(ProgramIdProperty))
            End Get
            Set(value As String)
                SetValue(ProgramIdProperty, value)
            End Set
        End Property

        Public Property Rating As Double
            Get
                Return GetValue(RatingProperty)
            End Get
            Set(value As Double)
                SetValue(RatingProperty, value)
            End Set
        End Property

        Public Property Stars As ObservableCollection(Of Polygon)
        Public Event RatingChanged As EventHandler

        Public Sub New()
            Stars = New ObservableCollection(Of Polygon)()
            Me.InitializeComponent()

            For i = 0 To 5 - 1
                Dim item As Polygon = New Polygon()
                Stars.Add(item)
            Next

            filledBrush = New SolidColorBrush(filledColor)
            unfilledBrush = New SolidColorBrush(unfilledColor)
        End Sub

        Protected Overrides Sub OnPreviewMouseDown(e As MouseButtonEventArgs)
            Rating = Math.Min(CInt(6.0 * e.GetPosition(Me).X / ActualWidth), 4) + 1
        End Sub

        Protected Overrides Sub OnPreviewMouseMove(e As MouseEventArgs)
            Dim num = Math.Min(CInt(6.0 * e.GetPosition(Me).X / ActualWidth), 4)

            For i = 0 To num + 1 - 1
                Stars(i).Fill = filledBrush
            Next

            For j = num + 1 To 5 - 1
                Stars(j).Fill = unfilledBrush
            Next
        End Sub

        Protected Overrides Sub OnMouseLeave(e As MouseEventArgs)
            RepaintRating()
        End Sub

        Private Shared Sub OnRatingChanged(obj As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim ratingControl As RatingControl = TryCast(obj, RatingControl)
            ratingControl.RepaintRating()
            ratingControl.FireRatingChanged(obj, EventArgs.Empty)
            Dim num As Double = e.OldValue
            Dim num2 As Double = e.NewValue

            If num <> num2 Then
                Dim service As Service = New Service()
                service.SubmitRatingAsync(ratingControl.ProgramId, num2)
            End If
        End Sub

        Private Sub FireRatingChanged(sender As Object, e As EventArgs)
            RaiseEvent RatingChanged(sender, e)
        End Sub

        Private Sub RepaintRating()
            If Not Rating < 1.0 AndAlso Not Rating > 5.0 AndAlso Not Double.IsNaN(Rating) Then
                Dim i = 0

                While i < Math.Floor(Rating)
                    Stars(i).Fill = filledBrush
                    i += 1
                End While

                For j = Math.Floor(Rating) To 5 - 1
                    Stars(j).Fill = unfilledBrush
                Next

                If Math.Floor(Rating) <> Rating Then
                    Dim num As Integer = Math.Floor(Rating)
                    Dim offset = Rating - num
                    Dim linearGradientBrush As LinearGradientBrush = New LinearGradientBrush(filledColor, unfilledColor, 0.0)
                    linearGradientBrush.GradientStops.Add(New GradientStop(filledColor, offset))
                    linearGradientBrush.GradientStops.Add(New GradientStop(unfilledColor, offset))
                    Stars(num).Fill = linearGradientBrush
                End If
            End If
        End Sub
    End Class
End Namespace
