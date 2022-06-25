Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Media

Namespace Microsoft.Nautilus.Text.Editor
    Public NotInheritable Class DefaultSpanGeometry
        Implements ISpanGeometry

        Private Const _roundness As Double = 0.0
        Private Const _overlap As Double = 0.0
        Private _textView As ITextView

        Public Sub New(textView As ITextView)
            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            _textView = textView
        End Sub

        Public Function GetMarkerGeometry(span1 As Span) As Geometry Implements ISpanGeometry.GetMarkerGeometry
            If span1.End > _textView.TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            If span1.IsEmpty Then
                Return Nothing
            End If

            Dim pathGeometry1 As New PathGeometry
            pathGeometry1.FillRule = FillRule.Nonzero
            Dim formattedTextLines1 As IFormattedTextLineCollection = _textView.FormattedTextLines

            If formattedTextLines1.Count = 0 Then Return Nothing

            If span1.End < formattedTextLines1(0).LineSpan.Start OrElse span1.Start > formattedTextLines1(formattedTextLines1.Count - 1).LineSpan.End Then
                Return Nothing
            End If

            Dim index As Integer
            formattedTextLines1.FindTextLineIndexContainingPosition(span1.Start, index)

            Dim textLine As ITextLine = formattedTextLines1(Math.Max(0, index))
            Dim textLine2 As ITextLine = Nothing
            Dim index2 As Integer

            formattedTextLines1.FindTextLineIndexContainingPosition(span1.End, index2)
            textLine2 = formattedTextLines1(Math.Max(0, index2))

            If textLine Is textLine2 Then
                Dim textBounds1 As ReadOnlyCollection(Of TextBounds) = textLine.GetTextBounds(span1)
                For Each item As TextBounds In textBounds1
                    Dim rectangleGeometry1 As New RectangleGeometry(New Rect(item.Left - 0.0, item.Top - 1.0, item.Width + 0.0, item.Height + 1.0), 0.0, 0.0)
                    rectangleGeometry1.Freeze()
                    pathGeometry1.AddGeometry(rectangleGeometry1)
                Next

            Else
                Dim textBounds2 As ReadOnlyCollection(Of TextBounds) = textLine.GetTextBounds(span1)
                For Each item2 As TextBounds In textBounds2
                    Dim rectangleGeometry2 As New RectangleGeometry(New Rect(item2.Left - 0.0, item2.Top - 1.0, item2.Width + 0.0, item2.Height + 1.0), 0.0, 0.0)
                    rectangleGeometry2.Freeze()
                    pathGeometry1.AddGeometry(rectangleGeometry2)
                Next

                Dim num As Integer = formattedTextLines1.IndexOf(textLine) + 1
                Dim num2 As Integer = formattedTextLines1.IndexOf(textLine2)
                For i As Integer = num To num2 - 1
                    Dim textLine3 As ITextLine = formattedTextLines1(i)
                    Dim rectangleGeometry3 As New RectangleGeometry(New Rect(textLine3.Left - 0.0, textLine3.Top - 0.0 - 1.0, textLine3.Width, textLine3.Height + 0.0 + 1.0), 0.0, 0.0)
                    rectangleGeometry3.Freeze()
                    pathGeometry1.AddGeometry(rectangleGeometry3)
                Next

                textBounds2 = textLine2.GetTextBounds(span1)
                For Each item3 As TextBounds In textBounds2
                    Dim rectangleGeometry4 As New RectangleGeometry(New Rect(item3.Left - 0.0, item3.Top - 1.0, item3.Width + 0.0, item3.Height + 1.0), 0.0, 0.0)
                    rectangleGeometry4.Freeze()
                    pathGeometry1.AddGeometry(rectangleGeometry4)
                Next
            End If

            pathGeometry1.Freeze()
            Return pathGeometry1.GetOutlinedPathGeometry()
        End Function
    End Class
End Namespace
