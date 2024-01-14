Public Class WndProperties

    Public Property LeftValue As Double?
        Get
            Return If(NumLeft.CheckBox.IsChecked,
                            GetValue(LeftValueProperty),
                            Nothing)
        End Get

        Set(value As Double?)
            SetDoubleValue(value, LeftValueProperty, NumLeft)
        End Set
    End Property

    Private Sub SetDoubleValue(doubleValue As Double?, dp As DependencyProperty, num As WpfDialogs.DoubleUpDown)
        If doubleValue Is Nothing OrElse Double.IsNaN(doubleValue) Then
            num.IsCheckable = True
            SetValue(dp, 0.0)
            num.CheckBox.IsChecked = False
        Else
            Dim chk = num.CheckBox
            If chk IsNot Nothing Then
                num.CheckBox.IsChecked = True
                num.IsCheckable = False
            End If
            SetValue(dp, doubleValue.Value)
        End If
    End Sub

    Public Shared ReadOnly LeftValueProperty As DependencyProperty =
                           DependencyProperty.Register("LeftValue",
                           GetType(Double), GetType(WndProperties),
                           New PropertyMetadata(Double.NaN))

    Public Property TopValue As Double?
        Get
            Return If(NumTop.CheckBox.IsChecked,
                            GetValue(TopValueProperty),
                            Nothing)
        End Get

        Set(value As Double?)
            SetDoubleValue(value, TopValueProperty, NumTop)
        End Set
    End Property

    Public Shared ReadOnly TopValueProperty As DependencyProperty =
                           DependencyProperty.Register("TopValue",
                           GetType(Double), GetType(WndProperties),
                           New PropertyMetadata(Double.NaN))

    Public Property WidthValue As Double?
        Get
            If Not NumWidth.CheckBox.IsChecked Then Return Nothing
            Dim v = CDbl(GetValue(WidthValueProperty))
            Return If(v <= 0, Double.NaN, v)
        End Get

        Set(value As Double?)
            SetDoubleValue(value, WidthValueProperty, NumWidth)
        End Set
    End Property

    Public Shared ReadOnly WidthValueProperty As DependencyProperty =
                           DependencyProperty.Register("WidthValue",
                           GetType(Double), GetType(WndProperties),
                           New PropertyMetadata(700.0))

    Public Property HeightValue As Double?
        Get
            If Not NumHeight.CheckBox.IsChecked Then Return Nothing
            Dim v = CDbl(GetValue(HeightValueProperty))
            Return If(v <= 0, Double.NaN, v)
        End Get

        Set(value As Double?)
            SetDoubleValue(value, HeightValueProperty, NumHeight)
        End Set
    End Property

    Public Shared ReadOnly HeightValueProperty As DependencyProperty =
                           DependencyProperty.Register("HeightValue",
                           GetType(Double), GetType(WndProperties),
                           New PropertyMetadata(500.0))


    Public Property MinWidthValue As Double?
        Get
            If Not NumMinWidth.CheckBox.IsChecked Then Return Nothing
            Return CDbl(GetValue(MinWidthValueProperty))
        End Get

        Set(value As Double?)
            SetDoubleValue(
                  value,
                  MinWidthValueProperty,
                  NumMinWidth
            )
        End Set
    End Property

    Public Shared ReadOnly MinWidthValueProperty As DependencyProperty =
                           DependencyProperty.Register("MinWidthValue",
                           GetType(Double), GetType(WndProperties),
                           New PropertyMetadata(0.0))

    Public Property MinHeightValue As Double?
        Get
            If Not NumMinHeight.CheckBox.IsChecked Then Return Nothing
            Return CDbl(GetValue(MinHeightValueProperty))
        End Get

        Set(value As Double?)
            SetDoubleValue(
                  value,
                  MinHeightValueProperty,
                  NumMinHeight
            )
        End Set
    End Property

    Public Shared ReadOnly MinHeightValueProperty As DependencyProperty =
                           DependencyProperty.Register("MinHeightValue",
                           GetType(Double), GetType(WndProperties),
                           New PropertyMetadata(0.0))

    Public Property MaxWidthValue As Double?
        Get
            If Not NumMaxWidth.CheckBox.IsChecked Then Return Nothing
            Dim v = CDbl(GetValue(MaxWidthValueProperty))
            Return If(v <= 0, Double.PositiveInfinity, v)
        End Get

        Set(value As Double?)
            SetDoubleValue(
                  If(value.HasValue AndAlso Double.IsPositiveInfinity(value), 0, value),
                  MaxWidthValueProperty,
                  NumMaxWidth
            )
        End Set
    End Property

    Public Shared ReadOnly MaxWidthValueProperty As DependencyProperty =
                           DependencyProperty.Register("MaxWidthValue",
                           GetType(Double), GetType(WndProperties),
                           New PropertyMetadata(0.0))

    Public Property MaxHeightValue As Double?
        Get
            If Not NumMaxHeight.CheckBox.IsChecked Then Return Nothing
            Dim v = CDbl(GetValue(MaxHeightValueProperty))
            Return If(v <= 0, Double.PositiveInfinity, v)
        End Get

        Set(value As Double?)
            SetDoubleValue(
                  If(value.HasValue AndAlso Double.IsPositiveInfinity(value), 0, value),
                  MaxHeightValueProperty,
                  NumMaxHeight
            )
        End Set
    End Property

    Public Shared ReadOnly MaxHeightValueProperty As DependencyProperty =
                           DependencyProperty.Register("MaxHeightValue",
                           GetType(Double), GetType(WndProperties),
                           New PropertyMetadata(0.0))


    Public Property EnabledValue As Boolean?
        Get
            Return IndexToBoolean(GetValue(EnabledValueProperty))
        End Get

        Set(value As Boolean?)
            SetValue(EnabledValueProperty, GetIndex(value))
        End Set
    End Property

    Private Function IndexToBoolean(index As Object) As Boolean?
        Dim id = CInt(index)
        Select Case id
            Case -1
                Return Nothing
            Case 0
                Return False
            Case Else
                Return True
        End Select
    End Function

    Private Function GetIndex(value As Boolean?) As Integer
        If Not value.HasValue Then Return -1
        If value.Value Then Return 1
        Return 0
    End Function

    Public Shared ReadOnly EnabledValueProperty As DependencyProperty =
                           DependencyProperty.Register("EnabledValue",
                           GetType(Integer), GetType(WndProperties),
                           New PropertyMetadata(1))


    Public Property VisibleValue As Boolean?
        Get
            Return IndexToBoolean(GetValue(VisibleValueProperty))
        End Get

        Set(value As Boolean?)
            SetValue(VisibleValueProperty, GetIndex(value))
        End Set
    End Property

    Public Shared ReadOnly VisibleValueProperty As DependencyProperty =
                           DependencyProperty.Register("VisibleValue",
                           GetType(Integer), GetType(WndProperties),
                           New PropertyMetadata(1))


    Public Property RightToLeftValue As Boolean?
        Get
            Return IndexToBoolean(GetValue(RightToLeftValueProperty))
        End Get

        Set(value As Boolean?)
            SetValue(RightToLeftValueProperty, GetIndex(value))
        End Set
    End Property

    Public Shared ReadOnly RightToLeftValueProperty As DependencyProperty =
                           DependencyProperty.Register("RightToLeftValue",
                           GetType(Integer), GetType(WndProperties),
                           New PropertyMetadata(0))


    Public Property WordWrapValue As Boolean?
        Get
            Return IndexToBoolean(GetValue(WordWrapValueProperty))
        End Get

        Set(value As Boolean?)
            SetValue(WordWrapValueProperty, GetIndex(value))
        End Set
    End Property

    Public Shared ReadOnly WordWrapValueProperty As DependencyProperty =
                           DependencyProperty.Register("WordWrapValue",
                           GetType(Integer), GetType(WndProperties),
                           New PropertyMetadata(-1))


    Public Property TagValue As String
        Get
            If Not chkTag.IsChecked Then Return Nothing
            Return GetValue(TagValueProperty)
        End Get

        Set(value As String)
            If value Is Nothing Then
                chkTag.Visibility = Visibility.Visible
                chkTag.IsChecked = False
                SetValue(TagValueProperty, "")
            Else
                chkTag.Visibility = Visibility.Collapsed
                chkTag.IsChecked = True
                SetValue(TagValueProperty, value)
            End If
        End Set
    End Property

    Public Shared ReadOnly TagValueProperty As DependencyProperty =
                           DependencyProperty.Register("TagValue",
                           GetType(String), GetType(WndProperties),
                           New PropertyMetadata(""))


    Public Property ToolTipValue As String
        Get
            If Not chkToolTip.IsChecked Then Return Nothing
            Return GetValue(ToolTipValueProperty)
        End Get

        Set(value As String)
            If value Is Nothing Then
                chkToolTip.Visibility = Visibility.Visible
                chkToolTip.IsChecked = False
                SetValue(ToolTipValueProperty, "")
            Else
                chkToolTip.Visibility = Visibility.Collapsed
                chkToolTip.IsChecked = True
                SetValue(ToolTipValueProperty, value)
            End If
        End Set
    End Property

    Public Shared ReadOnly ToolTipValueProperty As DependencyProperty =
                           DependencyProperty.Register("ToolTipValue",
                           GetType(String), GetType(WndProperties),
                           New PropertyMetadata(""))

    Private Sub BtnOk_Click(sender As Object, e As RoutedEventArgs)
        If NumMaxWidth.Value AndAlso NumMaxWidth.Value < NumMinWidth.Value Then
            NumMaxWidth.Focus()
            MsgBox("MaxWidth can't be less than MinWidth")
        ElseIf NumMaxHeight.Value > 0 AndAlso NumMaxHeight.Value < NumMinHeight.Value Then
            NumMaxHeight.Focus()
            MsgBox("MaxHeight can't be less than MinHeight")
        Else
            Me.DialogResult = True
        End If
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
    End Sub

    Private Sub WndProperties_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        NumLeft.Focus()
        NumLeft.TextBox.SelectAll()
    End Sub

    Private Sub chkTag_Checked(sender As Object, e As RoutedEventArgs) Handles chkTag.Checked
        If txtTag IsNot Nothing Then txtTag.IsEnabled = True
    End Sub

    Private Sub chkTag_Unchecked(sender As Object, e As RoutedEventArgs) Handles chkTag.Unchecked
        If txtTag IsNot Nothing Then txtTag.IsEnabled = False
    End Sub

    Private Sub chkToolTip_Checked(sender As Object, e As RoutedEventArgs) Handles chkToolTip.Checked
        If txtToolTip IsNot Nothing Then txtToolTip.IsEnabled = True
    End Sub

    Private Sub chkToolTip_Unchecked(sender As Object, e As RoutedEventArgs) Handles chkToolTip.Unchecked
        If txtToolTip IsNot Nothing Then txtToolTip.IsEnabled = False
    End Sub
End Class
