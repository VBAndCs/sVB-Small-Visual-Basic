
Imports System.Globalization
Imports System.Text

Namespace Library
    ''' <summary>
    ''' The primitive type representing either text or number.
    ''' </summary>
    Public Structure Primitive
        Private Class PrimitiveComparer
            Implements IEqualityComparer(Of Primitive)
            Private Shared _instance As New PrimitiveComparer

            Public Shared ReadOnly Property Instance As PrimitiveComparer
                Get
                    Return _instance
                End Get
            End Property

            Private Sub New()
            End Sub

            Public Function Equals(x As Primitive, y As Primitive) As Boolean Implements Collections.Generic.IEqualityComparer(Of Microsoft.SmallBasic.Library.Primitive).Equals
                Return String.Equals(x.AsString(), y.AsString(), StringComparison.InvariantCultureIgnoreCase)
            End Function

            Public Function GetHashCode(obj As Primitive) As Integer Implements Collections.Generic.IEqualityComparer(Of Microsoft.SmallBasic.Library.Primitive).GetHashCode
                Return obj.GetHashCode()
            End Function
        End Class

        Public Shared Function DateToPrimitive(ticks As Long) As Primitive
            Return New Primitive(ticks, NumberType.Date)
        End Function

        Public Shared Function TimeSpanToPrimitive(ticks As Long) As Primitive
            Return New Primitive(ticks, NumberType.TimeSpan)
        End Function

        Friend Function AsDate() As Date?
            If IsEmpty OrElse IsArray Then Return Nothing
            If _decimalValue.HasValue Then Return New Date(CLng(_decimalValue.Value))

            Dim d As Date
            If Date.TryParse(AsString(), d) Then
                Return d
            Else
                Return Nothing
            End If
        End Function

        Friend Function AsTimeSpan() As TimeSpan?
            If IsEmpty OrElse IsArray Then Return Nothing
            If _decimalValue.HasValue Then Return New TimeSpan(CLng(_decimalValue.Value))

            Dim ts As TimeSpan
            If TimeSpan.TryParse(AsString(), ts) Then
                Return ts
            Else
                Return Nothing
            End If
        End Function

        Private _stringValue As String
        Private _decimalValue As Decimal?

        Friend _arrayMap As Dictionary(Of Primitive, Primitive)
        Friend _isArray As Boolean

        Default Public Property Items(index As Primitive) As Primitive
            Get
                ConstructArrayMap()
                If Not IsArray Then
                    Dim s = If(_stringValue, "")
                    If index < 1 OrElse index > s.Length Then Return ""
                    Return s(index - 1).ToString()

                ElseIf CBool(ContainsKey(index)) Then
                    Return _arrayMap(index)
                End If

                Return ""
            End Get

            Set(Value As Primitive)
                ConstructArrayMap()
                If Not IsArray AndAlso _stringValue <> "" Then
                    If index < 1 Then
                        Me._stringValue = CStr(Value) & Space(-index) & _stringValue
                    ElseIf index > _stringValue.Length Then
                        Me._stringValue &= Space(index - _stringValue.Length - 1) & CStr(Value)
                    Else
                        Me._stringValue = _stringValue.Substring(0, index - 1) & CStr(Value) & _stringValue.Substring(index)
                    End If

                Else
                    _decimalValue = Nothing
                    _arrayMap(index) = Value
                End If
            End Set
        End Property

        Friend Function AsString() As String
            If _decimalValue.HasValue Then
                Select Case numberType
                    Case NumberType.Date
                        Dim d = New Date(_decimalValue.Value)
                        If d.Day = 1 AndAlso d.Month = 1 AndAlso d.Year = 1 Then
                            _stringValue = d.ToLongTimeString()
                        Else
                            _stringValue = d.ToString()
                        End If

                    Case NumberType.TimeSpan
                        _stringValue = New TimeSpan(_decimalValue.Value).ToString()
                    Case Else
                        _stringValue = _decimalValue.Value.ToString()
                End Select

            ElseIf _arrayMap IsNot Nothing AndAlso _arrayMap.Count > 0 Then
                Dim sb As New StringBuilder
                For Each entry In _arrayMap
                    sb.AppendFormat("{0}={1};", Escape(entry.Key), Escape(entry.Value))
                Next

                _stringValue = sb.ToString()
            End If

            Return If(_stringValue, "")
        End Function

        Friend ReadOnly Property IsArray As Boolean
            Get
                If _isArray Then Return True
                ConstructArrayMap()
                Return _arrayMap.Count > 0
            End Get
        End Property

        Friend ReadOnly Property IsEmpty As Boolean
            Get
                Return _stringValue = "" AndAlso
                        Not _decimalValue.HasValue AndAlso
                        (_arrayMap Is Nothing OrElse _arrayMap.Count = 0)
            End Get
        End Property

        Friend ReadOnly Property IsNumber As Boolean
            Get
                If _stringValue = "" Then Return True ' Consider it 0

                Dim result As Decimal = 0D
                If Not _decimalValue.HasValue Then
                    Return Decimal.TryParse(_stringValue, NumberStyles.Float, CultureInfo.InvariantCulture, result)
                End If

                Return True
            End Get
        End Property

        Public Sub New(primitiveText As String)
            _stringValue = If(primitiveText, "")
            _decimalValue = Nothing
            _arrayMap = Nothing
        End Sub

        Public Sub New(primitiveInteger As Integer)
            _decimalValue = primitiveInteger
            _stringValue = Nothing
            _arrayMap = Nothing
        End Sub

        Dim numberType As NumberType

        Public Sub New(primitiveDecimal As Decimal, Optional numberType As NumberType = NumberType.Decimal)
            _decimalValue = primitiveDecimal
            _stringValue = Nothing
            _arrayMap = Nothing
            Me.numberType = numberType
        End Sub

        Public Sub New(primitiveFloat As Single)
            Me.New(CDec(primitiveFloat))
        End Sub

        Public Sub New(primitiveDouble As Double)
            Me.New(CDec(primitiveDouble))
        End Sub

        Public Sub New(primitiveShort As Short)
            Me.New(CInt(primitiveShort))
        End Sub

        Public Sub New(primitiveLong As Long)
            _decimalValue = primitiveLong
            _stringValue = Nothing
            _arrayMap = Nothing
        End Sub

        Public Sub New(primitiveObject As Object)
            _stringValue = primitiveObject.ToString()
            _decimalValue = Nothing
            _arrayMap = Nothing
        End Sub

        Public Sub New(primitiveBool As Boolean)
            _decimalValue = Nothing
            If primitiveBool Then
                _stringValue = "True"
            Else
                _stringValue = "False"
            End If
            _arrayMap = Nothing
        End Sub

        Public Function Append(primitive As Primitive) As Primitive
            Return New Primitive(AsString() & primitive.AsString())
        End Function

        Public Function Add(addend As Primitive) As Primitive
            Dim n1 = TryGetAsDecimal()
            Dim n2 = addend.TryGetAsDecimal()
            If n1.HasValue AndAlso n2.HasValue Then
                Return New Primitive(
                        n1.Value + n2.Value,
                        GetNumberType(Me.numberType, addend.numberType)
                )
            End If
            Return New Primitive(AsString() & addend.AsString())
        End Function

        Public Function ContainsKey(key1 As Primitive) As Primitive
            ConstructArrayMap()
            Return _arrayMap.ContainsKey(key1)
        End Function

        Public Function ContainsValue(value As Primitive) As Primitive
            ConstructArrayMap()
            Return _arrayMap.ContainsValue(value)
        End Function

        Public Function GetAllIndices() As Primitive
            ConstructArrayMap()
            Dim indices As Dictionary(Of Primitive, Primitive)

            If IsArray Then
                indices = New Dictionary(Of Primitive, Primitive)(_arrayMap.Count, PrimitiveComparer.Instance)
                Dim key = 1
                For Each index In _arrayMap.Keys
                    indices(key) = index
                    key += 1
                Next
            Else
                ' Treat string as array of characters
                Dim count = AsString().Length
                indices = New Dictionary(Of Primitive, Primitive)(count, PrimitiveComparer.Instance)
                For index = 1 To count
                    indices(index) = index
                Next
            End If

            Return ConvertFromMap(indices)
        End Function

        Public Function GetItemCount() As Primitive
            ConstructArrayMap()
            Return _arrayMap.Count
        End Function

        Public Function Subtract(addend As Primitive) As Primitive
            Return New Primitive(
                        AsDecimal() - addend.AsDecimal(),
                        GetNumberType(Me.numberType, addend.numberType)
               )
        End Function

        Private Function GetNumberType(numberType1 As NumberType, numberType2 As NumberType) As NumberType

            Select Case numberType1
                Case NumberType.Decimal
                    Return NumberType.Decimal

                Case NumberType.Date
                    If numberType2 = NumberType.Date Then
                        Return NumberType.TimeSpan
                    Else
                        Return NumberType.Date
                    End If
                Case Else
                    Return NumberType.TimeSpan
            End Select
        End Function

        Public Function Multiply(multiplicand As Primitive) As Primitive
            Return New Primitive(AsDecimal() * multiplicand.AsDecimal())
        End Function

        Public Function Divide(divisor As Primitive) As Primitive
            divisor.AsDecimal()
            Return New Primitive(AsDecimal() / divisor.AsDecimal())
        End Function

        Friend ReadOnly Property IsTimeSpan As Boolean
            Get
                Return numberType = NumberType.TimeSpan
            End Get
        End Property

        Public Function LessThan(comparer As Primitive) As Boolean
            Return AsDecimal() < comparer.AsDecimal()
        End Function

        Friend ReadOnly Property IsDate As Boolean
            Get
                Return numberType = NumberType.Date
            End Get
        End Property

        Public Function GreaterThan(comparer As Primitive) As Boolean
            Return AsDecimal() > comparer.AsDecimal()
        End Function

        Public Function LessThanOrEqualTo(comparer As Primitive) As Boolean
            Return AsDecimal() <= comparer.AsDecimal()
        End Function

        Public Function GreaterThanOrEqualTo(comparer As Primitive) As Boolean
            Return AsDecimal() >= comparer.AsDecimal()
        End Function

        Public Function EqualTo(comparer As Primitive) As Boolean
            Return Equals(comparer)
        End Function

        Public Function NotEqualTo(comparer As Primitive) As Boolean
            Return Not Equals(comparer)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If AsString() Is Nothing Then _stringValue = ""

            If TypeOf obj Is Primitive Then
                Dim p1 = CType(obj, Primitive)
                Dim s1 = p1.AsString()
                Dim s2 = Me.AsString()
                Dim b1 = s1.ToLower(CultureInfo.InvariantCulture)
                Dim b2 = s2.ToLower(CultureInfo.InvariantCulture)

                If b1 = "true" OrElse b1 = "false" Then
                    Return CBool(Me) = Boolean.Parse(b1)

                ElseIf b2 = "true" OrElse b2 = "false" Then
                    Return CBool(p1) = Boolean.Parse(b2)

                ElseIf Me.IsNumber AndAlso p1.IsNumber Then
                    Return Me.AsDecimal() = p1.AsDecimal()

                Else
                    Return s1 = s2
                End If
            End If

            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            If AsString() Is Nothing Then _stringValue = ""

            Return AsString().ToUpper(CultureInfo.InvariantCulture).GetHashCode()
        End Function

        Public Overrides Function ToString() As String
            Return AsString()
        End Function

        Private Sub ConstructArrayMap()
            If _arrayMap?.Count > 0 Then Return

            _arrayMap = New Dictionary(Of Primitive, Primitive)(PrimitiveComparer.Instance)

            If IsEmpty Then Return

            Dim source = AsString().ToCharArray()

            Dim index As Integer = 0
            Dim L = source.Length
            Do
                Dim key = Unescape(source, index)
                If key = "" OrElse index = L Then
                    _arrayMap.Clear()
                    Return
                End If

                Dim value = Unescape(source, index)
                If value Is Nothing Then
                    _arrayMap.Clear()
                    Return
                End If
                _arrayMap(key) = value
            Loop While index < L
        End Sub

        Friend Function TryGetAsDecimal() As Decimal?
            If _decimalValue.HasValue Then Return _decimalValue

            If IsArray Then Return Nothing

            If _stringValue = "" Then Return 0D

            Dim result As Decimal
            If Decimal.TryParse(_stringValue, NumberStyles.Float, CultureInfo.InvariantCulture, result) Then
                _decimalValue = result
            End If

            Return _decimalValue
        End Function

        Friend Function AsDecimal() As Decimal
            If _decimalValue.HasValue Then
                Return _decimalValue.Value
            End If

            If _stringValue = "" Then Return 0D

            Dim result As Decimal = 0D
            If Decimal.TryParse(_stringValue, NumberStyles.Float, CultureInfo.InvariantCulture, result) Then
                _decimalValue = result
            End If

            Return result
        End Function

        Public Shared Function ConvertToBoolean(primitive As Primitive) As Boolean
            Return CBool(primitive)
        End Function

        Public Shared Widening Operator CType(primitive As Primitive) As String
            Return If(primitive.AsString(), "")
        End Operator

        Public Shared Widening Operator CType(primitive As Primitive) As Integer
            Return CInt(primitive.AsDecimal())
        End Operator

        Public Shared Widening Operator CType(primitive1 As Primitive) As Single
            Return CSng(primitive1.AsDecimal())
        End Operator

        Public Shared Widening Operator CType(primitive As Primitive) As Double
            Return CDbl(primitive.AsDecimal())
        End Operator

        Public Shared Widening Operator CType(primitive As Primitive) As Decimal
            Return primitive.AsDecimal()
        End Operator

        Public Shared Widening Operator CType(value As Primitive) As Boolean
            Dim s = value.AsString()
            If s = "" OrElse s = "0" Then Return False
            s = s.ToLower()
            If s = "false" Then Return False
            If s = "true" Then Return True
            If value.IsArray Then Return value._arrayMap?.Count > 0
            Dim d = value.TryGetAsDecimal()
            Return d.HasValue  ' we tested for zoro before, so, if it is numeric it is non zero, if it's not numeric, it is any string which means false
        End Operator

        Public Shared Operator =(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return primitive1.Equals(primitive2)
        End Operator

        Public Shared Operator <>(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return Not primitive1.Equals(primitive2)
        End Operator

        Public Shared Operator >(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return primitive1.GreaterThan(primitive2)
        End Operator

        Public Shared Operator >=(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return primitive1.GreaterThanOrEqualTo(primitive2)
        End Operator

        Public Shared Operator <(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return primitive1.LessThan(primitive2)
        End Operator

        Public Shared Operator <=(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            If primitive1.LessThanOrEqualTo(primitive2) Then
                Return New Primitive(True)
            Else
                Return New Primitive(False)
            End If
        End Operator

        Public Shared Operator +(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return primitive1.Add(primitive2)
        End Operator

        Public Shared Operator -(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return primitive1.Subtract(primitive2)
        End Operator

        Public Shared Operator *(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return primitive1.Multiply(primitive2)
        End Operator

        Public Shared Operator /(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return primitive1.Divide(primitive2)
        End Operator

        Public Shared Operator -(primitive1 As Primitive) As Primitive
            Return -primitive1.AsDecimal()
        End Operator

        Public Shared Operator Or(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return op_Or(primitive1, primitive2)
        End Operator

        Public Shared Operator And(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return op_And(primitive1, primitive2)
        End Operator

        Public Shared Operator IsTrue(primitive As Primitive) As Boolean
            Return CBool(primitive)
        End Operator

        Public Shared Operator IsFalse(primitive As Primitive) As Boolean
            Return Not CBool(primitive)
        End Operator

        Public Shared Function op_And(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return CBool(primitive1) AndAlso CBool(primitive2)
        End Function

        Public Shared Function op_Or(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return CBool(primitive1) OrElse CBool(primitive2)
        End Function

        Public Shared Widening Operator CType(value As Integer) As Primitive
            Return New Primitive(value)
        End Operator

        Public Shared Widening Operator CType(value As Boolean) As Primitive
            Return New Primitive(value)
        End Operator

        Public Shared Widening Operator CType(value As String) As Primitive
            If value Is Nothing Then
                Return New Primitive("")
            End If

            Dim p As New Primitive(value)
            If value.Contains("=") AndAlso value.Contains(";") Then
                p.ConstructArrayMap()
            End If
            Return p
        End Operator

        Public Shared Widening Operator CType(value As Double) As Primitive
            Return New Primitive(CDec(value))
        End Operator

        Public Shared Widening Operator CType(value As Decimal) As Primitive
            Return New Primitive(value)
        End Operator

        Public Shared Widening Operator CType(value As DateTime) As Primitive
            Return New Primitive(value)
        End Operator

        Public Shared Function GetArrayValue(array As Primitive, indexer As Primitive) As Primitive
            array.ConstructArrayMap()
            If array.GetItemCount() = 0 Then
                ' index the string
                Return array(indexer)
            Else
                Dim value As Primitive = Nothing
                If Not array._arrayMap.TryGetValue(indexer, value) Then
                    value = CType(Nothing, Primitive)
                End If
                Return value
            End If
        End Function

        Public Shared Function InitializeArray(value As Primitive) As Primitive
            Dim array As New Primitive
            array.ConstructArrayMap()
            Dim arrStr = CStr(value).Trim("{"c, "}"c, " "c)
            Dim arr = arrStr.Split({",", " "}, StringSplitOptions.RemoveEmptyEntries)
            array._arrayMap = New Dictionary(Of Primitive, Primitive)(PrimitiveComparer.Instance)

            For i = 0 To arr.Length - 1
                array._arrayMap.Add(CStr(i + 1), arr(i))
            Next

            Return array
        End Function


        Public Shared Function SetArrayValue(
                        value As Primitive,
                        array As Primitive,
                        indexer As Primitive
                   ) As Primitive

            array.ConstructArrayMap()
            If array.GetItemCount = 0 AndAlso CStr(array) <> "" Then
                ' imdex the string
                array(indexer) = value
                Return array
            Else
                Dim map As New Dictionary(Of Primitive, Primitive)(array._arrayMap, PrimitiveComparer.Instance)
                map(indexer) = value
                Return ConvertFromMap(map)
            End If

        End Function

        Public Shared Function ConvertFromMap(map As Dictionary(Of Primitive, Primitive)) As Primitive
            Dim result = CType(Nothing, Primitive)
            result._stringValue = Nothing
            result._decimalValue = Nothing
            result._arrayMap = map
            Return result
        End Function

        Private Shared Function Escape(value As String) As String
            Dim stringBuilder1 As New StringBuilder
            For Each c As Char In value
                Select Case c
                    Case "="c
                        stringBuilder1.Append("=")
                    Case ";"c
                        stringBuilder1.Append("\;")
                    Case "\"c
                        stringBuilder1.Append("\")
                    Case Else
                        stringBuilder1.Append(c)
                End Select
            Next

            Return stringBuilder1.ToString()
        End Function

        Private Shared Function Unescape(source As Char(), ByRef index As Integer) As String
            Dim flag = False
            Dim empty = True
            Dim length = source.Length
            Dim sb As New StringBuilder

            While index < length
                Dim c As Char = source(index)
                index += 1
                If Not flag Then
                    If c = "\"c Then
                        flag = True
                        Continue While
                    End If

                    If c = "="c Then
                        If empty Then Return Nothing
                        Return sb.ToString()
                    ElseIf c = ";"c Then
                        Return sb.ToString()
                    End If
                Else
                    flag = False
                End If

                empty = False
                sb.Append(c)
            End While

            Return Nothing
        End Function
    End Structure

    Public Enum NumberType
        [Decimal]
        [Date]
        [TimeSpan]
    End Enum

End Namespace
