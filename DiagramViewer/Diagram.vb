Imports System.Windows.Controls.Primitives
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Input

Public Class DiagramObject
    Dim WithEvents Diagram As FrameworkElement
    Dim Pnl As DiagramPanel
    Dim DesignerItem As ListBoxItem
    Dim Dsn As Designer
    Dim Canv As Canvas
    Dim Scv As ScrollViewer

    Friend Shared Diagrams As New Dictionary(Of FrameworkElement, DiagramObject)

    Private Sub New(Diagram As FrameworkElement)
        Me.Diagram = Diagram
        Canv = Helper.GetCanvas(Diagram)
        Scv = Helper.GetScrollViewer(Canv)
        Dsn = Helper.GetDesigner(Scv)
    End Sub

    Shared Function CreateDiagramObject(Diagram As FrameworkElement) As DiagramObject
        Dim DiagramObj As DiagramObject

        If Diagrams.ContainsKey(Diagram) Then
            DiagramObj = Diagrams(Diagram)
        Else
            DiagramObj = New DiagramObject(Diagram)
            Diagrams.Add(Diagram, DiagramObj)
        End If

        DiagramObj.Pnl = Helper.GetDiagramPanel(Diagram)
        DiagramObj.DesignerItem = Helper.GetListBoxItem(DiagramObj.Pnl)

        Return DiagramObj
    End Function

    Sub OutlineText()
        Dim Tb = Pnl.DiagramTextBlock
        Dim TextForeground = Designer.GetDiagramTextForeground(Diagram)
        Dim TextBackground = Designer.GetDiagramTextBackground(Diagram)

        Dim Gd1 As New GeometryDrawing(TextBackground, Nothing,
                                   New RectangleGeometry(New Rect(0, 0, Tb.ActualWidth, Tb.ActualHeight)))

        Dim Tf = New Typeface(Tb.FontFamily, Tb.FontStyle, Tb.FontWeight, Tb.FontStretch)
        Dim F As New FormattedText(Tb.Text, CultureInfo.CurrentCulture, Tb.FlowDirection,
                                   Tf, Tb.FontSize, TextForeground)

        Dim Gd2 As New GeometryDrawing(Designer.GetDiagramTextOutlineFill(Diagram), New Pen(TextForeground, Designer.GetDiagramTextOutlineThickness(Diagram)), F.BuildGeometry(New Point))
        Dim Dg As New DrawingGroup
        Dg.Children.Add(Gd1)
        Dg.Children.Add(Gd2)
        Dim TextBursh As New DrawingBrush(Dg)
        Tb.Background = TextBursh
        Tb.Foreground = Nothing
    End Sub

    Private Sub Diagram_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Diagram.PreviewMouseLeftButtonDown
        If e.ClickCount > 1 Then
            Dsn.OnDiagramDoubleClick(Diagram)
        End If
    End Sub

End Class
