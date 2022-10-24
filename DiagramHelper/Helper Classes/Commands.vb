Public Class Commands

    Shared Sub ChangeBackground(Element As FrameworkElement)
        If TypeOf Element Is Shape Then
            ChangeBrush(Element, Shape.FillProperty)
        Else
            ChangeBrush(Element, Control.BackgroundProperty)
        End If
    End Sub

    Shared Function ChangeBorderBrush(Element As FrameworkElement) As Brush
        Dim Shape As Shape = TryCast(Element, Shape)
        Dim B As Brush
        If Shape IsNot Nothing Then
            B = ChangeBrush(Element, Shape.StrokeProperty)
            If Shape.Stroke IsNot Nothing AndAlso (Double.IsNaN(Shape.StrokeThickness) OrElse Shape.StrokeThickness = 0) Then Shape.StrokeThickness = 1
        Else
            Dim Control = TryCast(Element, Control)
            If Control IsNot Nothing Then
                B = ChangeBrush(Element, Control.BorderBrushProperty)
                If Control.BorderBrush IsNot Nothing AndAlso (Double.IsNaN(Control.BorderThickness.Left) OrElse Control.BorderThickness.Left = 0) Then Control.BorderThickness = New Thickness(1)
            End If
        End If
        Return B
    End Function

    Shared Sub IncreaseBorderThickness(Element As FrameworkElement, Value As Double)
        Dim Shape As Shape = TryCast(Element, Shape)
        Dim Dsn = Helper.GetDesigner(Element)

        If Shape IsNot Nothing Then
            If Value < 0 AndAlso Shape.StrokeThickness = 0 Then Return
            Dim OldState As New PropertyState(AddressOf DiagramObject.Diagrams(Element).AfterRestoreAction, Element, Shape.StrokeThicknessProperty)
            Shape.StrokeThickness += Value
            Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
        Else
            Dim Control = TryCast(Element, Control)
            If Control IsNot Nothing Then
                Dim t = Control.BorderThickness.Left
                If Value < 0 AndAlso t = 0 Then Return
                Dim OldState As New PropertyState(AddressOf DiagramObject.Diagrams(Element).AfterRestoreAction, Element, Control.BorderThicknessProperty)
                Control.BorderThickness = New Thickness(t + Value)
                Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
            End If
        End If
    End Sub

    Shared Function ChangeBrush(Element As DependencyObject, Prop As DependencyProperty) As Brush


        If WpfDialogs.ColorDialog.Show(Element.GetValue(Prop)) Then
            Dim A As Action = Nothing
            If TypeOf Element Is FrameworkElement AndAlso DiagramObject.Diagrams.ContainsKey(Element) Then
                A = AddressOf DiagramObject.Diagrams(Element).AfterRestoreAction
            End If
            Dim OldState As New PropertyState(A, Element, Prop)
            Element.SetValue(Prop, WpfDialogs.ColorDialog.Brush)
            If TypeOf Element Is FrameworkElement Then
                If OldState.HasChanges Then
                    Dim Dsn = Helper.GetDesigner(Element)
                    Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
                End If
            End If
            Return WpfDialogs.ColorDialog.Brush
        End If
        Return Nothing
    End Function

    Shared Sub IncreaseRotationAngle(Element As FrameworkElement, Value As Double)
        Dim Angle As Double = Designer.GetRotationAngle(Element) + Value
        ChangeProperty(Element, Designer.RotationAngleProperty, Angle)
    End Sub

    Shared Sub ApplyLastChangeTo(Element As FrameworkElement)
        Dim Dsn = Helper.GetDesigner(Element)
        Dim Unit = Dsn.UndoStack.LastChange
        If Unit Is Nothing Then Return

        Dim PropState = TryCast(Unit(0), PropertyState)
        If PropState Is Nothing Then Return

        Dim OldState As New PropertyState(AddressOf DiagramObject.Diagrams(Element).AfterRestoreAction, Element, PropState.Keys.ToArray)
        For Each Pair In PropState
            Element.SetValue(Pair.Key, Helper.Clone(Pair.Value.NewValue))
        Next

        Dim Pnl = Helper.GetDiagramPanel(Element)
        If Pnl.DiagramGroup IsNot Nothing Then Pnl.DiagramGroup.Select()

        If OldState.HasChanges Then
            Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue), False)
        End If
    End Sub

    Shared Sub Skew(Element As FrameworkElement)
        Dim WndSkew As New WndSkew
        Dim T = TryCast(Element.LayoutTransform, SkewTransform)
        If T IsNot Nothing Then WndSkew.SkewTransform = New SkewTransform(T.AngleX, T.AngleY, Element.RenderTransformOrigin.X, Element.RenderTransformOrigin.Y)
        If WndSkew.ShowDialog() = True Then
            Dim OldState As New PropertyState(AddressOf DiagramObject.Diagrams(Element).AfterRestoreAction, Element, FrameworkElement.RenderTransformOriginProperty, FrameworkElement.LayoutTransformProperty)
            Dim ST = WndSkew.SkewTransform
            Element.RenderTransformOrigin = New Point(ST.CenterX, ST.CenterY)
            Dim Ax = ST.AngleX + If(ST.AngleX Mod 180 = 0, 0.000000000001, 0)
            Dim Ay = ST.AngleY + If(ST.AngleY Mod 180 = 0, 0.000000000001, 0)
            Element.LayoutTransform = New SkewTransform(Ax, Ay)

            Helper.UpdateControl(Element)
            Dim Pnl = Helper.GetDiagramPanel(Element)
            Pnl.DiagramGroup?.UpdateSelection()
            Pnl.UpdateLocationBorder()

            If OldState.HasChanges Then
                Pnl.Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
            End If
        End If
    End Sub

    Shared Sub ChangeTextBrush(Diagram As FrameworkElement, Prop As DependencyProperty)
        If WpfDialogs.ColorDialog.Show(Diagram.GetValue(Prop)) Then
            Dim OldState As New PropertyState(Diagram, Designer.DiagramTextFontPropsProperty)
            Diagram.SetValue(Prop, WpfDialogs.ColorDialog.Brush)
            If OldState.HasChanges Then
                Dim Dsn = Helper.GetDesigner(Diagram)
                Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
            End If
        End If
    End Sub

    Shared Sub ChangeFont(Diagram As FrameworkElement)
        If TypeOf Diagram IsNot Control Then Return

        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        Dim FontProps = Designer.GetDiagramTextFontProps(Diagram)
        FontProps.Add(WpfDialogs.FontDialog.FontProperties.ToArray)
        Dim OldState As New PropertyState(Diagram, Designer.DiagramTextFontPropsProperty)
        If WpfDialogs.FontDialog.Show(Diagram) Then
            FontProps.UpdateValuesFromObj()
            If OldState.HasChanges Then
                Pnl.Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
            End If
        End If
    End Sub

    Shared Sub IncreaseFontSize(Diagram As FrameworkElement, Value As Integer)
        Dim control = TryCast(Diagram, Control)
        If control Is Nothing Then Return

        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        Dim Sz = control.FontSize
        If (Value > 0 AndAlso Sz + Value <= 100) OrElse (Value < 0 AndAlso Sz + Value >= 1) Then
            ChangeFontProperty(Diagram, TextBlock.FontSizeProperty, Sz + Value)
        End If

    End Sub

    Shared Sub ChangeFontProperty(diagram As FrameworkElement, Prop As DependencyProperty, Value As Object)
        Dim Pnl = Helper.GetDiagramPanel(diagram)
        If Pnl Is Nothing Then Return

        Dim A As Action = AddressOf Pnl.DiagramObj.AfterRestoreAction
        Dim FontProps As New PropertyDictionary(diagram, Prop)
        Designer.SetDiagramTextFontProps(diagram, FontProps)
        Dim OldState As New PropertyState(A, diagram, Designer.DiagramTextFontPropsProperty)
        FontProps(Prop) = Value
        Pnl.Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))

    End Sub

    Shared Sub UpdateFontProperties(Diagram As Object)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        Dim FontProps = Designer.GetDiagramTextFontProps(Diagram)
        If FontProps Is Nothing Then Return
        FontProps.Add(WpfDialogs.FontDialog.FontProperties.ToArray)
        Designer.SetDiagramTextFontProps(Diagram, FontProps)
    End Sub


    Shared Sub ChangeProperty(Diagram As FrameworkElement, Prop As DependencyProperty, Value As Object)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        Dim A As Action = AddressOf Pnl.DiagramObj.AfterRestoreAction
        Dim OldState As New PropertyState(A, Diagram, Prop)
        Diagram.SetValue(Prop, Value)
        Pnl.Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
    End Sub
End Class
