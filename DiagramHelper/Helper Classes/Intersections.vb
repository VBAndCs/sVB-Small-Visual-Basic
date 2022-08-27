Friend Class Intersections

    Shared Function AreIntersected(R As Rect, PathPoints As List(Of Point)) As Boolean
        For i = 0 To PathPoints.Count - 1
            If R.Contains(PathPoints(i)) Then Return True
        Next

        For i = 0 To PathPoints.Count - 2
            If LinesIntersect(R.TopLeft, R.TopRight, PathPoints(i), PathPoints(i + 1)) Then Return True
            If LinesIntersect(R.TopLeft, R.BottomLeft, PathPoints(i), PathPoints(i + 1)) Then Return True
            If LinesIntersect(R.TopRight, R.BottomRight, PathPoints(i), PathPoints(i + 1)) Then Return True
            If LinesIntersect(R.BottomLeft, R.BottomRight, PathPoints(i), PathPoints(i + 1)) Then Return True
        Next

        Return False
    End Function

    Shared Function LinesIntersect(p1 As Point, p2 As Point, p3 As Point, p4 As Point) As Boolean
        Dim P As New Point(Double.NaN, Double.NaN)
        Dim a1 As Double = Double.NaN
        Dim a2 As Double = Double.NaN

        If p2.X = p1.X Then
            P.X = p1.X
        Else
            a1 = (p2.Y - p1.Y) / (p2.X - p1.X)
        End If

        If p4.X = p3.X Then
            P.X = p3.X
        Else
            a2 = (p4.Y - p3.Y) / (p4.X - p3.X)
        End If

        Dim b1 = p1.Y - a1 * p1.X
        Dim b2 = p3.Y - a2 * p3.X

        If Double.IsNaN(P.X) Then
            P.X = (b2 - b1) / (a1 - a2)
            P.Y = a2 * P.X + b2
            Return PointOnTheLines(P, p1, p2, p3, p4)

        ElseIf Not Double.IsNaN(a1) Then
            P.Y = a1 * P.X + b1
            Return PointOnTheLines(P, p1, p2, p3, p4)

        ElseIf Not Double.IsNaN(a2) Then
            P.Y = a2 * P.X + b2
            Return PointOnTheLines(P, p1, p2, p3, p4)

        ElseIf p1.X = p3.X Then
            Dim MinY = Math.Round(Math.Min(p3.Y, p4.Y), 2)
            If Math.Round(p1.Y, 2) < MinY AndAlso p2.Y < MinY Then Return False
            Dim MaxY = Math.Round(Math.Max(p3.Y, p4.Y), 2)
            If Math.Round(p1.Y, 2) > MaxY AndAlso p2.Y > MaxY Then Return False
            Return True
        Else
            Return False
        End If

    End Function

    Private Shared Function PointOnTheLines(P As Point, p1 As Point, p2 As Point, p3 As Point, p4 As Point) As Boolean
        P.X = Math.Round(P.X, 2)
        P.Y = Math.Round(P.Y, 2)
        Return P.X >= Math.Round(Math.Min(p1.X, p2.X), 2) AndAlso
                   P.Y >= Math.Round(Math.Min(p1.Y, p2.Y), 2) AndAlso
                   P.X <= Math.Round(Math.Max(p1.X, p2.X), 2) AndAlso
                   P.Y <= Math.Round(Math.Max(p1.Y, p2.Y), 2) AndAlso
                   P.X >= Math.Round(Math.Min(p3.X, p4.X), 2) AndAlso
                   P.Y >= Math.Round(Math.Min(p3.Y, p4.Y), 2) AndAlso
                   P.X <= Math.Round(Math.Max(p3.X, p4.X), 2) AndAlso
                   P.Y <= Math.Round(Math.Max(p3.Y, p4.Y), 2)
    End Function

End Class
