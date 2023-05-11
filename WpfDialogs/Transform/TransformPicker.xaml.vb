Public Class TransformPicker

    Public Property Transform As Transform
        Get
            Return GetValue(TransformProperty)
        End Get

        Set(ByVal value As Transform)
            SetValue(TransformProperty, value)
        End Set
    End Property

    Public Shared ReadOnly TransformProperty As DependencyProperty = _
                           DependencyProperty.Register("Transform", _
                           GetType(Transform), GetType(TransformPicker), _
                           New PropertyMetadata(AddressOf OnTransformChanged))

    Public Event TransformChanged(Tp As TransformPicker, OldTransform As Transform, NewTransform As Transform)

    Friend Sub RaiseTransformChanged(Tp As TransformPicker, OldTransform As Transform, NewTransform As Transform)
        RaiseEvent TransformChanged(Tp, OldTransform, NewTransform)
    End Sub

    Dim DontChangeTransform As Boolean = False

    Shared Sub OnTransformChanged(Tp As TransformPicker, e As DependencyPropertyChangedEventArgs)
        If Tp.DontChangeTransform Then Return

        Dim Trans As Transform = e.NewValue
        If Trans Is Tp.TransformGroup Then Return

        Tp.DontChangeTransform = True

        Tp.chkScale.IsChecked = False
        Tp.chkRotate.IsChecked = False
        Tp.chkSkew.IsChecked = False
        Tp.chkTrnanslate.IsChecked = False
        Tp.chkMatrix.IsChecked = False

        If TypeOf Trans Is TransformGroup Then
            Dim Tg As TransformGroup = Trans
            For i = 0 To Tg.Children.Count - 1
                Tp.SetTransform(Tg.Children(i), i)
            Next
            Tp.ModifyTranformGroup()           
        Else
            Tp.SetTransform(Trans, 0)
        End If

        Tp.DontChangeTransform = False
    End Sub


    Private Sub SetTransform(Trans As Transform, Index As Integer)
        If Trans Is Nothing Then Return
        Select Case Trans.GetType
            Case GetType(ScaleTransform)
                Me.ScaleTransform = Trans
                Me.chkScale.IsChecked = True

            Case GetType(RotateTransform)
                Me.RotateTransform = Trans
                Me.chkRotate.IsChecked = True

            Case GetType(SkewTransform)
                Me.SkewTransform = Trans
                Me.chkSkew.IsChecked = True

            Case GetType(TranslateTransform)
                Me.TranslateTransform = Trans
                Me.chkTrnanslate.IsChecked = True

            Case GetType(MatrixTransform)
                If CType(Trans, MatrixTransform).Matrix = Matrix.Identity Then Return
                Me.MatrixTransform = Trans
                Me.chkMatrix.IsChecked = True                
        End Select

        Dim g As Grid = LstTransform.SelectedItem
        Dim I = LstTransform.SelectedIndex
        If Index = I Then Return
        LstTransform.Items.RemoveAt(I)
        LstTransform.Items.Insert(Index, g)
        LstTransform.SelectedIndex = Index

    End Sub

    Public Property TransformGroup As TransformGroup
        Get
            Return GetValue(TransformGroupProperty)
        End Get

        Set(ByVal value As TransformGroup)
            SetValue(TransformGroupProperty, value)
        End Set
    End Property

    Public Shared ReadOnly TransformGroupProperty As DependencyProperty = _
                           DependencyProperty.Register("TransformGroup", _
                           GetType(TransformGroup), GetType(TransformPicker), New PropertyMetadata(AddressOf TransformGroupChanged))

    Shared Sub TransformGroupChanged(Tp As TransformPicker, e As DependencyPropertyChangedEventArgs)
        If Tp.DontChangeTransform Then Return
        Tp.Transform = e.NewValue
        Tp.RaiseTransformChanged(Tp, e.OldValue, e.NewValue)
    End Sub

    Public Property TargetWidth As Double
        Get
            Return GetValue(TargetWidthProperty)
        End Get

        Set(ByVal value As Double)
            SetValue(TargetWidthProperty, value)
        End Set
    End Property

    Public Shared ReadOnly TargetWidthProperty As DependencyProperty = _
                       DependencyProperty.Register("TargetWidth", _
                       GetType(Double), GetType(TransformPicker), _
                       New PropertyMetadata(1.0, AddressOf TargetWidthChanged))

    Shared Sub TargetWidthChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Tp As TransformPicker = d
        If Tp.Resources("RelativeXConverter") Is Nothing Then Tp.Resources("RelativeXConverter") = New RelativeConverter
        Dim XConverter As RelativeConverter = Tp.Resources("RelativeXConverter")
        XConverter.RelativeTo = e.NewValue
    End Sub

    Public Property TargetHeight As Double
        Get
            Return GetValue(TargetHeightProperty)
        End Get

        Set(ByVal value As Double)
            SetValue(TargetHeightProperty, value)
        End Set
    End Property

    Public Shared ReadOnly TargetHeightProperty As DependencyProperty = _
                           DependencyProperty.Register("TargetHeight", _
                           GetType(Double), GetType(TransformPicker), _
                           New PropertyMetadata(1.0, AddressOf TargetHeightChanged))

    Shared Sub TargetHeightChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Tp As TransformPicker = d
        If Tp.Resources("RelativeYConverter") Is Nothing Then Tp.Resources("RelativeYConverter") = New RelativeConverter
        Dim YConverter As RelativeConverter = Tp.Resources("RelativeYConverter")
        YConverter.RelativeTo = e.NewValue
    End Sub

    Public Property MatrixTransform As MatrixTransform
        Get
            Return GetValue(MatrixProperty)
        End Get

        Set(ByVal value As MatrixTransform)
            SetValue(MatrixProperty, value)
        End Set
    End Property

    Public Shared ReadOnly MatrixProperty As DependencyProperty = _
                           DependencyProperty.Register("MatrixTransform", _
                           GetType(MatrixTransform), GetType(TransformPicker), New PropertyMetadata(AddressOf MatrixTransformChanged))

    Shared Sub MatrixTransformChanged(Tp As TransformPicker, e As DependencyPropertyChangedEventArgs)
        Dim M = CType(e.NewValue, MatrixTransform).Matrix
        Tp.UdM11.Value = M.M11
        Tp.UdM12.Value = M.M12
        Tp.UdM21.Value = M.M21
        Tp.UdM22.Value = M.M22
        Tp.UdOffsetX.Value = M.OffsetX
        Tp.UdOffsetY.Value = M.OffsetY
    End Sub


    Public Property RotateTransform As RotateTransform
        Get
            Return GetValue(RotateTransformProperty)
        End Get

        Set(ByVal value As RotateTransform)
            SetValue(RotateTransformProperty, value)
        End Set
    End Property

    Public Shared ReadOnly RotateTransformProperty As DependencyProperty = _
                           DependencyProperty.Register("RotateTransform", _
                           GetType(RotateTransform), GetType(TransformPicker))


    Public Property ScaleTransform As ScaleTransform
        Get
            Return GetValue(ScaleTransformProperty)
        End Get

        Set(ByVal value As ScaleTransform)
            SetValue(ScaleTransformProperty, value)
        End Set
    End Property

    Public Shared ReadOnly ScaleTransformProperty As DependencyProperty = _
                           DependencyProperty.Register("ScaleTransform", _
                           GetType(ScaleTransform), GetType(TransformPicker))

    Public Property SkewTransform As SkewTransform
        Get
            Return GetValue(SkewTransformProperty)
        End Get

        Set(ByVal value As SkewTransform)
            SetValue(SkewTransformProperty, value)
        End Set
    End Property

    Public Shared ReadOnly SkewTransformProperty As DependencyProperty = _
                           DependencyProperty.Register("SkewTransform", _
                           GetType(SkewTransform), GetType(TransformPicker))


    Public Property TranslateTransform As TranslateTransform
        Get
            Return GetValue(TranslateTransformProperty)
        End Get

        Set(ByVal value As TranslateTransform)
            SetValue(TranslateTransformProperty, value)
        End Set
    End Property

    Public Shared ReadOnly TranslateTransformProperty As DependencyProperty = _
                           DependencyProperty.Register("TranslateTransform", _
                           GetType(TranslateTransform), GetType(TransformPicker))


    Private Sub BtnMoveUp_Click(sender As Object, e As RoutedEventArgs)
        Dim g As Grid = LstTransform.SelectedItem
        Dim I = LstTransform.SelectedIndex
        LstTransform.Items.RemoveAt(I)
        LstTransform.Items.Insert(I - 1, g)
        LstTransform.SelectedIndex = I - 1
        LstTransform.ScrollIntoView(g)
        ModifyTranformGroup()
    End Sub

    Private Sub BtnMoveDown_Click(sender As Object, e As RoutedEventArgs)
        Dim g As Grid = LstTransform.SelectedItem
        Dim I = LstTransform.SelectedIndex
        LstTransform.Items.RemoveAt(I)
        LstTransform.Items.Insert(I + 1, g)
        LstTransform.SelectedIndex = I + 1
        LstTransform.ScrollIntoView(g)
        ModifyTranformGroup()
    End Sub

    Sub ModifyTranformGroup()
        Dim Tg As New TransformGroup
        For Each g As Grid In LstTransform.Items
            Dim Chk As CheckBox = g.Children(1)
            If Chk.IsChecked Then
                Dim Exp As Expander = g.Children(0)
                Tg.Children.Add(Exp.DataContext)
            End If
        Next
        Me.TransformGroup = Tg
    End Sub

    Private Sub LstTransform_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        BtnMoveUp.IsEnabled = (LstTransform.SelectedIndex > 0)
        LblMoveUp.Background.Opacity = If(BtnMoveUp.IsEnabled, 1, 0.5)
        BtnMoveDown.IsEnabled = (LstTransform.SelectedIndex > -1 AndAlso LstTransform.SelectedIndex < LstTransform.Items.Count - 1)
        LblMoveDown.Background.Opacity = If(BtnMoveDown.IsEnabled, 1, 0.5)

        For Each Item In LstTransform.Items
            Dim G As Grid = Item
            Dim Chk As CheckBox = G.Children(1)
            Chk.Tag = If(Item Is LstTransform.SelectedItem, "IsSelected", "")
        Next

    End Sub

    Private Sub CheckBox_Checked(Chk As CheckBox, e As RoutedEventArgs)
        Dim G As Grid = Chk.Parent
        LstTransform.SelectedItem = G
        Dim Exp As Expander = G.Children(0)
        Exp.IsEnabled = True
        Exp.IsExpanded = True
        ModifyTranformGroup()
    End Sub

    Private Sub CheckBox_Unchecked(Chk As CheckBox, e As RoutedEventArgs)
        Dim G As Grid = Chk.Parent
        Dim Exp As Expander = G.Children(0)
        Exp.IsEnabled = False
        Exp.IsExpanded = False
        Me.TransformGroup.Children.Remove(Exp.DataContext)
    End Sub

    Private Sub Grid_GotKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        LstTransform.SelectedItem = sender
    End Sub

    Private Sub TransformPicker_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        DontChangeTransform = True
        If Me.TransformGroup Is Nothing Then Me.TransformGroup = New TransformGroup()
        If Me.ScaleTransform Is Nothing Then Me.ScaleTransform = New ScaleTransform
        If Me.RotateTransform Is Nothing Then Me.RotateTransform = New RotateTransform()
        If Me.SkewTransform Is Nothing Then Me.SkewTransform = New SkewTransform()
        If Me.TranslateTransform Is Nothing Then Me.TranslateTransform = New TranslateTransform()
        If Me.MatrixTransform Is Nothing Then Me.MatrixTransform = New MatrixTransform()
        DontChangeTransform = False
    End Sub

    Private Sub UdM11_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of System.Object))
        Dim M = Me.MatrixTransform.Matrix
        M.M11 = UdM11.Value
        Me.MatrixTransform.Matrix = M
    End Sub

    Private Sub UdM12_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of System.Object))
        Dim M = Me.MatrixTransform.Matrix
        M.M12 = UdM12.Value
        Me.MatrixTransform.Matrix = M
    End Sub

    Private Sub UdM21_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of System.Object))
        Dim M = Me.MatrixTransform.Matrix
        M.M21 = UdM21.Value
        Me.MatrixTransform.Matrix = M
    End Sub

    Private Sub UdM22_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of System.Object))
        Dim M = Me.MatrixTransform.Matrix
        M.M22 = UdM22.Value
        Me.MatrixTransform.Matrix = M
    End Sub

    Private Sub UdOffsetX_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of System.Object))
        Dim M = Me.MatrixTransform.Matrix
        M.OffsetX = UdOffsetX.Value
        Me.MatrixTransform.Matrix = M
    End Sub

    Private Sub UdOffsetY_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of System.Object))
        Dim M = Me.MatrixTransform.Matrix
        M.OffsetY = UdOffsetY.Value
        Me.MatrixTransform.Matrix = M
    End Sub
End Class
