Imports System.Globalization

<TemplatePart(Name:="PART_TextBox", Type:=GetType(TextBox))>
<TemplatePart(Name:="PART_Spinner", Type:=GetType(Spinner))>
<TemplatePart(Name:="PART_CheckBox", Type:=GetType(CheckBox))>
Public MustInherit Class UpDownBase
    Inherits Control
    Implements IValidateInput

#Region "Members"

    Friend Const PART_TextBox As String = "PART_TextBox"
    Friend Const PART_Spinner As String = "PART_Spinner"
    Friend Const PART_CheckBox As String = "PART_CheckBox"

    Private _isSyncingTextAndValueProperties As Boolean
    Private _isTextChangedFromUI As Boolean

#End Region


#Region "Properties"

    Friend Property Spinner() As Spinner
    Public ReadOnly Property TextBox() As TextBox

    Dim _CheckBox As CheckBox
    Public ReadOnly Property CheckBox() As CheckBox
        Get
            If _CheckBox IsNot Nothing Then Return _CheckBox
            Return GetCheckBox
        End Get
    End Property

    Private Function GetCheckBox() As CheckBox
        _CheckBox = TryCast(GetTemplateChild(PART_CheckBox), CheckBox)
        If _CheckBox IsNot Nothing Then
            _CheckBox.Visibility = If(IsCheckable, Visibility.Visible, Visibility.Collapsed)
            AddHandler _CheckBox.Checked, AddressOf OnChecked
            AddHandler _CheckBox.Unchecked, AddressOf OnUnchecked
        End If
        Return _CheckBox
    End Function

    Public Property IsCheckable As Boolean
        Get
            Return GetValue(IsCheckableProperty)
        End Get

        Set(ByVal value As Boolean)
            SetValue(IsCheckableProperty, value)
        End Set
    End Property

    Public Shared ReadOnly IsCheckableProperty As DependencyProperty =
                           DependencyProperty.Register("IsCheckable",
                           GetType(Boolean), GetType(UpDownBase),
                           New PropertyMetadata(False, AddressOf OnIsCheckableChanged))

    Private Shared Sub OnIsCheckableChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim ud = CType(d, UpDownBase)
        If ud.CheckBox Is Nothing Then Return
        ud.CheckBox.Visibility = If(e.NewValue, Visibility.Visible, Visibility.Collapsed)
    End Sub







#Region "CultureInfo"

    Public Shared ReadOnly CultureInfoProperty As DependencyProperty = DependencyProperty.Register("CultureInfo", GetType(CultureInfo), GetType(UpDownBase), New UIPropertyMetadata(CultureInfo.CurrentUICulture, AddressOf OnCultureInfoChanged))
    Public Property CultureInfo() As CultureInfo
        Get
            Return CType(GetValue(CultureInfoProperty), CultureInfo)
        End Get
        Set(ByVal value As CultureInfo)
            SetValue(CultureInfoProperty, value)
        End Set
    End Property

    Private Shared Sub OnCultureInfoChanged(ByVal o As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim inputBase As UpDownBase = TryCast(o, UpDownBase)
        If inputBase IsNot Nothing Then
            inputBase.OnCultureInfoChanged(CType(e.OldValue, CultureInfo), CType(e.NewValue, CultureInfo))
        End If
    End Sub

#End Region 'CultureInfo

#Region "IsReadOnly"

    Public Shared ReadOnly IsReadOnlyProperty As DependencyProperty = DependencyProperty.Register("IsReadOnly", GetType(Boolean), GetType(UpDownBase), New UIPropertyMetadata(False, AddressOf OnReadOnlyChanged))
    Public Property IsReadOnly() As Boolean
        Get
            Return CBool(GetValue(IsReadOnlyProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(IsReadOnlyProperty, value)
        End Set
    End Property

    Private Shared Sub OnReadOnlyChanged(ByVal o As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim inputBase As UpDownBase = TryCast(o, UpDownBase)
        If inputBase IsNot Nothing Then
            inputBase.OnReadOnlyChanged(CBool(e.OldValue), CBool(e.NewValue))
        End If
    End Sub

#End Region 'IsReadOnly

#Region "Text"

    Public Shared ReadOnly TextProperty As DependencyProperty = DependencyProperty.Register("Text",
                                              GetType(String), GetType(UpDownBase),
                                              New FrameworkPropertyMetadata("0.0", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
                                              AddressOf OnTextChanged, Nothing, False, UpdateSourceTrigger.LostFocus))

    Public Property Text() As String
        Get
            Return CStr(GetValue(TextProperty))
        End Get
        Set(ByVal value As String)
            SetValue(TextProperty, value)
        End Set
    End Property

    Private Shared Sub OnTextChanged(ByVal o As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim inputBase As UpDownBase = TryCast(o, UpDownBase)
        If inputBase IsNot Nothing Then
            inputBase.OnTextChanged(CStr(e.OldValue), CStr(e.NewValue))
        End If
    End Sub

#End Region 'Text

#Region "FormatString"

    Public Shared ReadOnly FormatStringProperty As DependencyProperty = DependencyProperty.Register("FormatString", GetType(String), GetType(UpDownBase), New UIPropertyMetadata(String.Empty, AddressOf OnFormatStringChanged))
    Public Property FormatString() As String
        Get
            Return CStr(GetValue(FormatStringProperty))
        End Get
        Set(ByVal value As String)
            SetValue(FormatStringProperty, value)
        End Set
    End Property

    Private Shared Sub OnFormatStringChanged(ByVal o As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim numericUpDown As UpDownBase = TryCast(o, UpDownBase)
        If numericUpDown IsNot Nothing Then
            numericUpDown.OnFormatStringChanged(CStr(e.OldValue), CStr(e.NewValue))
        End If
    End Sub

    Protected Overridable Sub OnFormatStringChanged(ByVal oldValue As String, ByVal newValue As String)
        If IsInitialized Then
            Me.SyncTextAndValue(False, Nothing)
        End If
    End Sub

#End Region 'FormatString

#Region "Increment"

    Public Shared ReadOnly IncrementProperty As DependencyProperty = DependencyProperty.Register("Increment", GetType(Double?), GetType(UpDownBase), New PropertyMetadata(Nothing, AddressOf OnIncrementChanged, AddressOf OnCoerceIncrement))
    Public Property Increment() As Double?
        Get
            Return CType(GetValue(IncrementProperty), Double?)
        End Get
        Set(ByVal value As Double?)
            SetValue(IncrementProperty, value)
        End Set
    End Property

    Private Shared Sub OnIncrementChanged(ByVal o As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim numericUpDown As UpDownBase = TryCast(o, UpDownBase)
        If numericUpDown IsNot Nothing Then
            numericUpDown.OnIncrementChanged(CDbl(e.OldValue), CDbl(e.NewValue))
        End If
    End Sub

    Protected Overridable Sub OnIncrementChanged(ByVal oldValue As Double, ByVal newValue As Double)
        If Me.IsInitialized Then
            SetValidSpinDirection()
        End If
    End Sub

    Private Shared Function OnCoerceIncrement(ByVal d As DependencyObject, ByVal baseValue As Object) As Object
        Dim numericUpDown As UpDownBase = TryCast(d, UpDownBase)
        If numericUpDown IsNot Nothing Then
            Return numericUpDown.OnCoerceIncrement(CDbl(baseValue))
        End If

        Return baseValue
    End Function

    Protected Overridable Function OnCoerceIncrement(ByVal baseValue? As Double) As Double?
        Return baseValue
    End Function

#End Region

#Region "Maximum"

    Public Shared ReadOnly MaximumProperty As DependencyProperty = DependencyProperty.Register("Maximum", GetType(Double), GetType(UpDownBase), New UIPropertyMetadata(Double.NaN, AddressOf OnMaximumChanged, AddressOf OnCoerceMaximum))
    Public Property Maximum() As Double
        Get
            Return CDbl(GetValue(MaximumProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(MaximumProperty, value)
        End Set
    End Property

    Private Shared Sub OnMaximumChanged(ByVal o As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim numericUpDown As UpDownBase = TryCast(o, UpDownBase)
        If numericUpDown IsNot Nothing Then
            numericUpDown.OnMaximumChanged(CDbl(e.OldValue), CDbl(e.NewValue))
        End If
    End Sub

    Protected Overridable Sub OnMaximumChanged(ByVal oldValue As Double, ByVal newValue As Double)
        If Me.IsInitialized Then
            SetValidSpinDirection()
        End If
    End Sub

    Private Shared Function OnCoerceMaximum(ByVal d As DependencyObject, ByVal baseValue As Object) As Object
        Dim numericUpDown As UpDownBase = TryCast(d, UpDownBase)
        If numericUpDown IsNot Nothing Then
            Return numericUpDown.OnCoerceMaximum(CDbl(baseValue))
        End If

        Return baseValue
    End Function

    Protected Overridable Function OnCoerceMaximum(ByVal baseValue As Double) As Double
        Return baseValue
    End Function

#End Region 'Maximum

#Region "Minimum"

    Public Shared ReadOnly MinimumProperty As DependencyProperty = DependencyProperty.Register("Minimum", GetType(Double), GetType(UpDownBase), New UIPropertyMetadata(Double.NaN, AddressOf OnMinimumChanged, AddressOf OnCoerceMinimum))
    Public Property Minimum() As Double
        Get
            Return CDbl(GetValue(MinimumProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(MinimumProperty, value)
        End Set
    End Property

    Private Shared Sub OnMinimumChanged(ByVal o As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim numericUpDown As UpDownBase = TryCast(o, UpDownBase)
        If numericUpDown IsNot Nothing Then
            numericUpDown.OnMinimumChanged(CDbl(e.OldValue), CDbl(e.NewValue))
        End If
    End Sub

    Protected Overridable Sub OnMinimumChanged(ByVal oldValue As Double, ByVal newValue As Double)
        If Me.IsInitialized Then
            SetValidSpinDirection()
        End If
    End Sub

    Private Shared Function OnCoerceMinimum(ByVal d As DependencyObject, ByVal baseValue As Object) As Object
        Dim numericUpDown As UpDownBase = TryCast(d, UpDownBase)
        If numericUpDown IsNot Nothing Then
            Return numericUpDown.OnCoerceMinimum(CDbl(baseValue))
        End If

        Return baseValue
    End Function

    Protected Overridable Function OnCoerceMinimum(ByVal baseValue As Double) As Double?
        Return baseValue
    End Function

#End Region 'Minimum

#Region "AllowSpin"

    Public Shared ReadOnly AllowSpinProperty As DependencyProperty = DependencyProperty.Register("AllowSpin", GetType(Boolean), GetType(UpDownBase), New UIPropertyMetadata(True))
    Public Property AllowSpin() As Boolean
        Get
            Return CBool(GetValue(AllowSpinProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(AllowSpinProperty, value)
        End Set
    End Property

#End Region 'AllowSpin

#Region "DefaultValue"

    Public Shared ReadOnly DefaultValueProperty As DependencyProperty = DependencyProperty.Register("DefaultValue", GetType(Double?), GetType(UpDownBase), New UIPropertyMetadata(Nothing, AddressOf OnDefaultValueChanged))
    Public Property DefaultValue() As Double?
        Get
            Return CType(GetValue(DefaultValueProperty), Double?)
        End Get
        Set(ByVal value As Double?)
            SetValue(DefaultValueProperty, value)
        End Set
    End Property

    Private Shared Sub OnDefaultValueChanged(ByVal source As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        CType(source, UpDownBase).OnDefaultValueChanged(CDbl(args.OldValue), CDbl(args.NewValue))
    End Sub

    Private Sub OnDefaultValueChanged(ByVal oldValue As Double, ByVal newValue As Double)
        If Me.IsInitialized AndAlso String.IsNullOrEmpty(Text) Then
            Me.SyncTextAndValue(True, Text)
        End If
    End Sub

#End Region 'DefaultValue

#Region "AllowInputSpecialValues"

    Private Shared ReadOnly AllowInputSpecialValuesProperty As DependencyProperty = DependencyProperty.Register("AllowInputSpecialValues", GetType(AllowedSpecialValues), GetType(UpDownBase), New UIPropertyMetadata(AllowedSpecialValues.None))

    Private Property AllowInputSpecialValues() As AllowedSpecialValues
        Get
            Return CType(GetValue(AllowInputSpecialValuesProperty), AllowedSpecialValues)
        End Get
        Set(ByVal value As AllowedSpecialValues)
            SetValue(AllowInputSpecialValuesProperty, value)
        End Set
    End Property

#End Region 'AllowInputSpecialValues

#Region "ParsingNumberStyle"

    Public Shared ReadOnly ParsingNumberStyleProperty As DependencyProperty = DependencyProperty.Register("ParsingNumberStyle", GetType(NumberStyles), GetType(UpDownBase), New UIPropertyMetadata(NumberStyles.Any))

    Public Property ParsingNumberStyle() As NumberStyles
        Get
            Return CType(GetValue(ParsingNumberStyleProperty), NumberStyles)
        End Get
        Set(ByVal value As NumberStyles)
            SetValue(ParsingNumberStyleProperty, value)
        End Set
    End Property

#End Region

#Region "Value"

    Public Shared ReadOnly ValueProperty As DependencyProperty = DependencyProperty.Register("Value", GetType(Double?), GetType(UpDownBase), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnValueChanged, AddressOf OnCoerceValue, False, UpdateSourceTrigger.PropertyChanged))
    Public Property Value() As Double?
        Get
            Return CType(GetValue(ValueProperty), Double?)
        End Get
        Set(ByVal value As Double?)
            Dim LastValue As Double? = GetValue(ValueProperty)
            If LastValue.Equals(value) Then Return
            SetValue(ValueProperty, value)
        End Set
    End Property

    Private Shared Function OnCoerceValue(ByVal o As DependencyObject, ByVal basevalue As Object) As Object
        Return CType(o, UpDownBase).OnCoerceValue(basevalue)
    End Function

    Protected Overridable Function OnCoerceValue(ByVal newValue As Object) As Object
        Return newValue
    End Function

    Private Shared Sub OnValueChanged(ByVal o As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim _upDownBase As UpDownBase = TryCast(o, UpDownBase)
        If _upDownBase IsNot Nothing Then
            _upDownBase.OnValueChanged(CDbl(e.OldValue), CDbl(e.NewValue))
        End If
    End Sub

    Protected Overridable Sub OnValueChanged(ByVal oldValue As Double, ByVal newValue As Double)
        If Me.IsInitialized Then
            SyncTextAndValue(False, Nothing)
        End If

        SetValidSpinDirection()

        Dim args As New RoutedPropertyChangedEventArgs(Of Object)(oldValue, newValue)
        args.RoutedEvent = ValueChangedEvent
        MyBase.RaiseEvent(args)
    End Sub

#End Region 'Value

#End Region 'Properties


#Region "Base Class Overrides"

    Protected Overrides Sub OnAccessKey(ByVal e As AccessKeyEventArgs)
        If TextBox IsNot Nothing Then
            TextBox.Focus()
        End If

        MyBase.OnAccessKey(e)
    End Sub

    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()

        _TextBox = TryCast(GetTemplateChild(PART_TextBox), TextBox)
        If TextBox IsNot Nothing Then
            If String.IsNullOrEmpty(Text) Then
                TextBox.Text = "0.0"
            Else
                TextBox.Text = Text
            End If

            AddHandler TextBox.LostFocus, AddressOf TextBox_LostFocus
            AddHandler TextBox.TextChanged, AddressOf TextBox_TextChanged
        End If

        If Spinner IsNot Nothing Then
            RemoveHandler Spinner.Spin, AddressOf OnSpinnerSpin
        End If

        Spinner = TryCast(GetTemplateChild(PART_Spinner), Spinner)

        If Spinner IsNot Nothing AndAlso AllowSpin AndAlso Not IsReadOnly Then
            AddHandler Spinner.Spin, AddressOf OnSpinnerSpin
        End If

        SetValidSpinDirection()

        GetCheckBox()
    End Sub

    Protected Overrides Sub OnGotFocus(ByVal e As RoutedEventArgs)
        If TextBox IsNot Nothing Then
            TextBox.Focus()
        End If
    End Sub

    Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)
        Select Case e.Key
            Case Key.Enter
                Dim commitSuccess As Boolean = CommitInput()
                e.Handled = Not commitSuccess
                Exit Select
        End Select
    End Sub

    Protected Sub OnTextChanged(ByVal oldValue As String, ByVal newValue As String)
        If Me.IsInitialized Then
            SyncTextAndValue(True, Text)
        End If
    End Sub

    Protected Sub OnCultureInfoChanged(ByVal oldValue As CultureInfo, ByVal newValue As CultureInfo)
        If IsInitialized Then
            SyncTextAndValue(False, Nothing)
        End If
    End Sub

    Protected Sub OnReadOnlyChanged(ByVal oldValue As Boolean, ByVal newValue As Boolean)
        SetValidSpinDirection()
    End Sub

    Protected Sub OnSpin(ByVal e As SpinEventArgs)
        If e Is Nothing Then
            Throw New ArgumentNullException("e")
        End If

        If e.Direction = SpinDirection.Increase Then
            OnIncrement()
        Else
            OnDecrement()
        End If
    End Sub

#End Region


#Region "Event Handlers"

    Private Sub OnSpinnerSpin(ByVal sender As Object, ByVal e As SpinEventArgs)
        If AllowSpin AndAlso (Not IsReadOnly) Then
            OnSpin(e)
        End If
    End Sub


    Private Sub OnUnchecked(sender As Object, e As RoutedEventArgs)
        Spinner.IsEnabled = False
    End Sub

    Private Sub OnChecked(sender As Object, e As RoutedEventArgs)
        Spinner.IsEnabled = True
    End Sub

#End Region


#Region "Events"

    Public Event InputValidationError As InputValidationErrorEventHandler Implements IValidateInput.InputValidationError

#Region "ValueChanged Event"

    Public Shared ReadOnly ValueChangedEvent As RoutedEvent = EventManager.RegisterRoutedEvent("ValueChanged", RoutingStrategy.Bubble, GetType(RoutedPropertyChangedEventHandler(Of Object)), GetType(UpDownBase))
    Public Custom Event ValueChanged As RoutedPropertyChangedEventHandler(Of Object)
        AddHandler(ByVal value As RoutedPropertyChangedEventHandler(Of Object))
            MyBase.AddHandler(ValueChangedEvent, value)
        End AddHandler
        RemoveHandler(ByVal value As RoutedPropertyChangedEventHandler(Of Object))
            MyBase.RemoveHandler(ValueChangedEvent, value)
        End RemoveHandler

        RaiseEvent(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object))
            MyBase.RaiseEvent(e)
        End RaiseEvent
    End Event



#End Region

#End Region 'Events


#Region "Methods"

    Private Sub TextBox_TextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs)
        Try
            _isTextChangedFromUI = True
            Text = CType(sender, TextBox).Text
        Finally
            _isTextChangedFromUI = False
        End Try
    End Sub

    Private Sub TextBox_LostFocus(ByVal sender As Object, ByVal e As RoutedEventArgs)
        CommitInput()
    End Sub

    Private Sub RaiseInputValidationError(ByVal e As Exception)
        If InputValidationErrorEvent IsNot Nothing Then
            Dim args As New InputValidationErrorEventArgs(e)
            RaiseEvent InputValidationError(Me, args)
            If args.ThrowException Then
                Throw args.Exception
            End If
        End If
    End Sub

    Public Function CommitInput() As Boolean Implements IValidateInput.CommitInput
        Return Me.SyncTextAndValue(True, Text)
    End Function

    Protected Function SyncTextAndValue(ByVal updateValueFromText As Boolean, ByVal text As String) As Boolean
        If _isSyncingTextAndValueProperties Then
            Return True
        End If

        _isSyncingTextAndValueProperties = True
        Dim parsedTextIsValid As Boolean = True
        Try
            If updateValueFromText Then
                If String.IsNullOrEmpty(text) Then
                    ' An empty input sets the value to the default value.
                    Value = Me.DefaultValue
                Else
                    Try
                        Value = Me.ConvertTextToValue(text)
                    Catch e As Exception
                        parsedTextIsValid = False

                        ' From the UI, just allow any input.
                        If Not _isTextChangedFromUI Then
                            ' This call may throw an exception. 
                            ' See RaiseInputValidationError() implementation.
                            Me.RaiseInputValidationError(e)
                        End If
                    End Try
                End If
            End If

            ' Do not touch the ongoing text input from user.
            If Not _isTextChangedFromUI Then
                ' Don't replace the empty Text with the non-empty representation of DefaultValue.
                Dim shouldKeepEmpty As Boolean = String.IsNullOrEmpty(Me.Text) AndAlso Object.Equals(Value, DefaultValue)
                If Not shouldKeepEmpty Then
                    Me.Text = ConvertValueToText()
                End If

                ' Sync Text and textBox
                If TextBox IsNot Nothing Then
                    If String.IsNullOrEmpty(Me.Text) Then
                        TextBox.Text = "0.0"
                    Else
                        TextBox.Text = Me.Text
                    End If
                End If
            End If

            If _isTextChangedFromUI AndAlso (Not parsedTextIsValid) Then
                '// Text input was made from the user and the text
                '// repesents an invalid value. Disable the spinner
                '// in this case.
                If Spinner IsNot Nothing Then
                    Spinner.ValidSpinDirection = ValidSpinDirections.None
                End If
            Else
                Me.SetValidSpinDirection()
            End If
        Finally
            _isSyncingTextAndValueProperties = False
        End Try
        Return parsedTextIsValid
    End Function

    Protected Shared Function ParsePercent(ByVal text As String, ByVal cultureInfo As IFormatProvider) As Decimal
        Dim info As NumberFormatInfo = NumberFormatInfo.GetInstance(cultureInfo)

        text = text.Replace(info.PercentSymbol, Nothing)

        Dim result As Decimal = Decimal.Parse(text, NumberStyles.Any, info)
        result = result / 100

        Return result
    End Function

#End Region


#Region "Abstract"

    Protected MustOverride Function ConvertTextToValue(ByVal text As String) As Double?
    Protected MustOverride Function ConvertValueToText() As String
    Protected MustOverride Sub OnIncrement()
    Protected MustOverride Sub OnDecrement()
    Protected MustOverride Sub SetValidSpinDirection()

#End Region

    Private Sub UpDownBase_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Me.SyncTextAndValue(False, Nothing)
    End Sub
End Class
