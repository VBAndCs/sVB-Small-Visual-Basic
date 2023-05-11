Imports System.Globalization
Imports System.IO

Public Class DoubleUpDown
    Inherits UpDownBase

    Protected Delegate Function FromText(ByVal s As String, ByVal style As NumberStyles, ByVal provider As IFormatProvider) As Double
    Protected Delegate Function FromDecimal(ByVal d As Decimal) As Double

    Private _fromText As FromText
    Private _fromDecimal As FromDecimal
    Private _fromLowerThan As Func(Of Double, Double, Boolean)
    Private _fromGreaterThan As Func(Of Double, Double, Boolean)

    Protected Sub New(ByVal fromText As FromText, ByVal fromDecimal As FromDecimal, ByVal fromLowerThan As Func(Of Double, Double, Boolean), ByVal fromGreaterThan As Func(Of Double, Double, Boolean))
        If fromText Is Nothing Then
            Throw New ArgumentNullException("parseMethod")
        End If

        If fromDecimal Is Nothing Then
            Throw New ArgumentNullException("fromDecimal")
        End If

        If fromLowerThan Is Nothing Then
            Throw New ArgumentNullException("fromLowerThan")
        End If

        If fromGreaterThan Is Nothing Then
            Throw New ArgumentNullException("fromGreaterThan")
        End If

        _fromText = fromText
        _fromDecimal = fromDecimal
        _fromLowerThan = fromLowerThan
        _fromGreaterThan = fromGreaterThan
    End Sub

    Protected Shared Sub UpdateMetadata(ByVal type As Type, ByVal increment? As Double, ByVal minValue? As Double, ByVal maxValue? As Double)
        DefaultStyleKeyProperty.OverrideMetadata(type, New FrameworkPropertyMetadata(type))
        IncrementProperty.OverrideMetadata(type, New FrameworkPropertyMetadata(increment))
        MaximumProperty.OverrideMetadata(type, New FrameworkPropertyMetadata(maxValue))
        MinimumProperty.OverrideMetadata(type, New FrameworkPropertyMetadata(minValue))
    End Sub

    Private Function IsLowerThan(ByVal value1? As Double, ByVal value2? As Double) As Boolean
        If value1 Is Nothing OrElse value2 Is Nothing Then
            Return False
        End If

        Return _fromLowerThan(value1.Value, value2.Value)
    End Function

    Private Function IsGreaterThan(ByVal value1? As Double, ByVal value2? As Double) As Boolean
        If value1 Is Nothing OrElse value2 Is Nothing Then
            Return False
        End If

        Return _fromGreaterThan(value1.Value, value2.Value)
    End Function

    Private Function HandleNullSpin() As Boolean
        If Not Value.HasValue Then
            Dim forcedValue As Double = If(DefaultValue.HasValue, DefaultValue.Value, Nothing)

            Value = CoerceValueMinMax(forcedValue)

            Return True
        ElseIf Not Increment.HasValue Then
            Return True
        End If

        Return False
    End Function

    Private Function CoerceValueMinMax(ByVal value As Double) As Double
        If IsLowerThan(value, Minimum) Then
            Return Minimum
        ElseIf IsGreaterThan(value, Maximum) Then
            Return Maximum
        Else
            Return value
        End If
    End Function

    Private Sub ValidateDefaultMinMax(ByVal value? As Double)
        If Object.Equals(value, DefaultValue) Then
            Return
        End If

        If IsLowerThan(value, Minimum) Then
            Throw New ArgumentOutOfRangeException("Minimum", String.Format("Value must be greater than MinValue of {0}", Minimum))
        ElseIf IsGreaterThan(value, Maximum) Then
            Throw New ArgumentOutOfRangeException("Maximum", String.Format("Value must be less than MaxValue of {0}", Maximum))
        End If
    End Sub

#Region "Base Class Overrides"

    Protected Overrides Sub OnIncrement()
        If Not HandleNullSpin() Then
            Dim result? As Double = Value.Value + Increment.Value
            Value = CoerceValueMinMax(result.Value)
        End If
    End Sub

    Protected Overrides Sub OnDecrement()
        If Not HandleNullSpin() Then
            Dim result? As Double = Value.Value - Increment.Value
            Value = CoerceValueMinMax(result.Value)
        End If
    End Sub

    Protected Overrides Function ConvertValueToText() As String
        If Value Is Nothing Then
            Return String.Empty
        End If

        Return Value.Value.ToString(FormatString, CultureInfo)
    End Function

    Protected Overrides Function ConvertTextToValue(ByVal text As String) As Double?
        Dim result? As Double = 0

        If String.IsNullOrEmpty(text) Then
            Return result
        End If

        Dim currentValueText As String = ConvertValueToText()
        If Object.Equals(currentValueText, text) Then
            Return Me.Value
        End If

        result = If(FormatString.Contains("P"), _fromDecimal(ParsePercent(text, CultureInfo)), _fromText(text, Me.ParsingNumberStyle, CultureInfo))

        ValidateDefaultMinMax(result)

        Return result
    End Function

    Protected Overrides Sub SetValidSpinDirection()
        Dim validDirections As ValidSpinDirections = ValidSpinDirections.None

        If (Me.Increment IsNot Nothing) AndAlso (Not IsReadOnly) Then
            If IsLowerThan(Value, Maximum) OrElse (Not Value.HasValue) Then
                validDirections = validDirections Or ValidSpinDirections.Increase
            End If

            If IsGreaterThan(Value, Minimum) OrElse (Not Value.HasValue) Then
                validDirections = validDirections Or ValidSpinDirections.Decrease
            End If
        End If

        If Spinner IsNot Nothing Then
            Spinner.ValidSpinDirection = validDirections
        End If
    End Sub

#End Region

#Region "Constructors"

    Shared Sub New()
        UpdateMetadata(GetType(DoubleUpDown), 1.0R, Double.NegativeInfinity, Double.PositiveInfinity)
    End Sub

    Public Sub New()
        Me.New(AddressOf Double.Parse, AddressOf Decimal.ToDouble, Function(v1, v2) v1 < v2, Function(v1, v2) v1 > v2)

    End Sub

#End Region
End Class

