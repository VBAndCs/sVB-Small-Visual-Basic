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

        Public Function GetMarkerGeometry(span As Span) As Geometry Implements ISpanGeometry.GetMarkerGeometry
            If span.IsEmpty Then Return Nothing

            If span.End > _textView.TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            Dim pathGeometry As New PathGeometry
            pathGeometry.FillRule = FillRule.Nonzero
            Dim formattedTextLines = _textView.FormattedTextLines
            If formattedTextLines.Count = 0 Then Return Nothing

            If span.End < formattedTextLines(0).LineSpan.Start OrElse span.Start > formattedTextLines(formattedTextLines.Count - 1).LineSpan.End Then
                Return Nothing
            End If

            Dim index As Integer
            formattedTextLines.FindTextLineIndexContainingPosition(span.Start, index)

            Dim textLine1 = formattedTextLines(Math.Max(0, index))
            Dim textLine2 As ITextLine = Nothing
            Dim index2 As Integer

            formattedTextLines.FindTextLineIndexContainingPosition(span.End, index2)
            textLine2 = formattedTextLines(Math.Max(0, index2))

            If textLine1 Is textLine2 Then
                For Each bounds In textLine1.GetTextBounds(span)
                    Dim rectangleGeometry As New RectangleGeometry(
                        New Rect(bounds.Left,
                                 bounds.Top - 1.0,
                                 bounds.Width,
                                 bounds.Height + 1.0),
                        0.0, 0.0)
                    rectangleGeometry.Freeze()
                    pathGeometry.AddGeometry(rectangleGeometry)
                Next

            Else
                For Each bounds In textLine1.GetTextBounds(span)
                    Dim rectangleGeometry As New RectangleGeometry(New Rect(bounds.Left - 0.0, bounds.Top - 1.0, bounds.Width + 0.0, bounds.Height + 1.0), 0.0, 0.0)
                    rectangleGeometry.Freeze()
                    pathGeometry.AddGeometry(rectangleGeometry)
                Next

                Dim lineIndex1 = formattedTextLines.IndexOf(textLine1) + 1
                Dim lineIndex2 = formattedTextLines.IndexOf(textLine2)

                For i = lineIndex1 To lineIndex2 - 1
                    Dim textLine3 As ITextLine = formattedTextLines(i)
                    Dim rectangleGeometry3 As New RectangleGeometry(New Rect(textLine3.Left - 0.0, textLine3.Top - 0.0 - 1.0, textLine3.Width, textLine3.Height + 0.0 + 1.0), 0.0, 0.0)
                    rectangleGeometry3.Freeze()
                    pathGeometry.AddGeometry(rectangleGeometry3)
                Next

                For Each bounds In textLine2.GetTextBounds(span)
                    Dim rectangleGeometry4 As New RectangleGeometry(New Rect(bounds.Left - 0.0, bounds.Top - 1.0, bounds.Width + 0.0, bounds.Height + 1.0), 0.0, 0.0)
                    rectangleGeometry4.Freeze()
                    pathGeometry.AddGeometry(rectangleGeometry4)
                Next
            End If

            pathGeometry.Freeze()
            Return pathGeometry.GetOutlinedPathGeometry()
        End Function
    End Class
End Namespace
