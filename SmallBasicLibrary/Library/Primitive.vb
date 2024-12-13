
Imports System.Globalization
Imports System.Text

Namespace Library
    ''' <summary>
    ''' The primitive type representing either text or number.
    ''' </summary>
    Public Structure Primitive
        Friend Class PrimitiveComparer
            Implements IEqualityComparer(Of Primitive)
            Private Shared _comparer As New PrimitiveComparer

            Public Shared ReadOnly Property Instance As PrimitiveComparer
                Get
                    Return _comparer
                End Get
            End Property

            Private Sub New()
            End Sub

            Public Function Equals(x As Primitive, y As Primitive) As Boolean Implements Collections.Generic.IEqualityComparer(Of Microsoft.SmallVisualBasic.Library.Primitive).Equals
                Return String.Equals(x.AsString(), y.AsString(), StringComparison.InvariantCultureIgnoreCase)
            End Function

            Public Function GetHashCode(obj As Primitive) As Integer Implements Collections.Generic.IEqualityComparer(Of Microsoft.SmallVisualBasic.Library.Primitive).GetHashCode
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
            If _decimalValue.HasValue Then
                Dim v = _decimalValue.Value
                If v < 0 Then Return Nothing
                Return New Date(v)
            End If

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

        Friend _stringValue As String
        Private _decimalValue As Decimal?

        Private _arrMap As Dictionary(Of Primitive, Primitive)
        Friend _isArray As Boolean

        Public Property ArrayMap As Dictionary(Of Primitive, Primitive)
            Get
                If (_arrMap Is Nothing OrElse _arrMap.Count = 0) AndAlso _stringValue <> "" Then
                    ConstructArrayMap()
                End If
                Return _arrMap
            End Get

            Set(value As Dictionary(Of Primitive, Primitive))
                _arrMap = value
            End Set
        End Property


        Default Public Property Items(index As Primitive) As Primitive
            Get
                ConstructArrayMap()
                If Not IsArray Then
                    If Not index.IsNumber Then
                        Return New Primitive("")
                    End If
                    Dim s = If(_stringValue, "")
                    If index < 1 OrElse index > s.Length Then Return New Primitive("")
                    Return New Primitive(s(index - 1).ToString())

                ElseIf CBool(ContainsKey(index)) Then
                    Return ArrayMap(index)
                End If

                Return New Primitive("")
            End Get

            Set(Value As Primitive)
                If IsArray OrElse _stringValue = "" Then
                    Me._decimalValue = Nothing
                    ArrayMap(index) = Value

                ElseIf index.IsNumber Then
                    If index < 1 Then
                        Me._stringValue = CStr(Value) & Space(-index) & _stringValue
                    ElseIf index > _stringValue.Length Then
                        Me._stringValue &= Space(index - _stringValue.Length - 1) & CStr(Value)
                    Else
                        Me._stringValue = _stringValue.Substring(0, index - 1) & CStr(Value) & _stringValue.Substring(index)
                    End If

                    If IsNumeric(Me._stringValue) Then
                        Me._decimalValue = Me._stringValue
                    Else
                        Me._decimalValue = Nothing
                    End If
                End If
            End Set
        End Property

        Friend Function AsString() As String
            If _decimalValue.HasValue Then
                Select Case numberType
                    Case NumberType.Date
                        Dim d = New Date(_decimalValue.Value)
                        If d.Day = 1 AndAlso d.Month = 1 AndAlso d.Year = 1 Then
                            _stringValue = d.ToString("T", New CultureInfo("en-Us"))
                        Else
                            _stringValue = d.ToString(
                                "MM/dd/yyyy hh:mm:ss.FFFFFFF tt",
                                New CultureInfo("en-US")
                            )
                        End If

                    Case NumberType.TimeSpan
                        Dim ts = New TimeSpan(_decimalValue.Value).ToString()
                        Dim pos = ts.LastIndexOf(":")
                        If pos = -1 Then
                            _stringValue = ts
                            Exit Select
                        End If

                        pos = ts.IndexOf(".", pos)
                        If pos = -1 Then
                            _stringValue = ts
                            Exit Select
                        End If

                        ts = ts.TrimEnd("0"c)
                        If ts.EndsWith(".") Then ts = ts.Substring(0, ts.Length - 1)
                        _stringValue = ts

                    Case Else
                        _stringValue = If(_stringValue <> "", _stringValue, _decimalValue.Value.ToString())
                        If _stringValue.Contains(".") Then
                            _stringValue = _stringValue.TrimEnd("0"c)
                            If _stringValue.EndsWith(".") Then
                                _stringValue = _stringValue.TrimEnd("."c)
                            End If
                        End If
                End Select

            ElseIf _arrMap IsNot Nothing AndAlso _arrMap.Count > 0 Then
                Dim sb As New StringBuilder
                For Each entry In ArrayMap
                    sb.AppendFormat("{0}={1};", Escape(entry.Key), Escape(entry.Value))
                Next

                _stringValue = sb.ToString()
            End If

            Return If(_stringValue, "")
        End Function

        Public ReadOnly Property IsBoolean As Boolean
            Get
                Dim s = AsString().ToLower()
                Return s = "false" OrElse s = "true"
            End Get
        End Property

        Public ReadOnly Property IsArray As Boolean
            Get
                If _isArray Then Return True
                ConstructArrayMap()
                Return ArrayMap.Count > 0
            End Get
        End Property

        Public ReadOnly Property IsEmpty As Boolean
            Get
                Return _stringValue = "" AndAlso
                        Not _decimalValue.HasValue AndAlso
                        (ArrayMap Is Nothing OrElse ArrayMap.Count = 0)
            End Get
        End Property

        Public ReadOnly Property IsNumber As Boolean
            Get
                If ArrayMap?.Count > 0 Then Return False
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
            _arrMap = Nothing
        End Sub

        Public Sub New(primitiveInteger As Integer)
            _decimalValue = primitiveInteger
            _stringValue = Nothing
            _arrMap = Nothing
        End Sub

        Dim numberType As NumberType

        Public Sub New(primitiveDecimal As Decimal, Optional numberType As NumberType = NumberType.Decimal)
            _decimalValue = primitiveDecimal
            _stringValue = Nothing
            _arrMap = Nothing
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
            ArrayMap = Nothing
        End Sub

        Public Sub New(primitiveObject As Object)
            _stringValue = primitiveObject.ToString()
            _decimalValue = Nothing
            ArrayMap = Nothing
        End Sub

        Public Sub New(arr() As String, removeEmpty As Boolean)
            ArrayMap = New Dictionary(Of Primitive, Primitive)
            _isArray = True

            For i = 0 To arr.Count - 1
                Dim item = arr(i)
                If Not removeEmpty OrElse item <> "" Then
                    ArrayMap(i + 1) = New Primitive(item)
                End If
            Next
        End Sub

        Public Sub New(primitiveBool As Boolean)
            _decimalValue = Nothing
            If primitiveBool Then
                _stringValue = "True"
            Else
                _stringValue = "False"
            End If
            ArrayMap = Nothing
        End Sub

        Public Function Append(primitive As Primitive) As Primitive
            Return New Primitive(AsString() & primitive.AsString())
        End Function

        Public Function Concat(secondPrimitive As Primitive) As Primitive
            Return New Primitive(Me.AsString() & secondPrimitive.AsString())
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

        Public Function ContainsKey(key As Primitive) As Primitive
            ConstructArrayMap()
            Return ArrayMap.ContainsKey(key)
        End Function

        Public Function ContainsValue(value As Primitive) As Primitive
            ConstructArrayMap()
            Return ArrayMap.ContainsValue(value)
        End Function

        Public Function GetAllIndices() As Primitive
            ConstructArrayMap()
            Dim indices As Dictionary(Of Primitive, Primitive)

            If IsArray Then
                indices = New Dictionary(Of Primitive, Primitive)(ArrayMap.Count, PrimitiveComparer.Instance)
                Dim key = 1
                For Each index In ArrayMap.Keys
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
            Return ArrayMap.Count
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
            If Me.IsNumber Then
                If multiplicand.IsNumber Then
                    Return New Primitive(Me.AsDecimal() * multiplicand.AsDecimal())
                ElseIf Not Me.IsArray Then
                    Return Duplicate(multiplicand.AsString(), Me.AsDecimal())
                End If
            ElseIf multiplicand.IsNumber AndAlso Not Me.IsArray Then
                Return Duplicate(Me.AsString(), multiplicand.AsDecimal())
            End If

            Return New Primitive("")
        End Function

        Friend Shared Function Duplicate(str As String, count As Integer) As Primitive
            If count = 0 Then Return New Primitive("")
            If count = 1 Then Return New Primitive(str)

            If str.Length = 1 Then
                Return New Primitive(New String(str(0), count))
            End If

            Dim sb As New StringBuilder
            For i = 1 To count
                sb.Append(str)
            Next
            Return New Primitive(sb.ToString())
        End Function

        Public Function Divide(divisor As Primitive) As Primitive
            Return New Primitive(AsDecimal() / divisor.AsDecimal())
        End Function
        Public Function Remainder(divisor As Primitive) As Primitive
            Return New Primitive(Decimal.Remainder(AsDecimal(), divisor.AsDecimal()))
        End Function

        Public ReadOnly Property IsTimeSpan As Boolean
            Get
                Return numberType = NumberType.TimeSpan
            End Get
        End Property

        Public Function LessThan(comparer As Primitive) As Boolean
            If Me.IsNumber AndAlso comparer.IsNumber Then
                Return Me.AsDecimal() < comparer.AsDecimal()
            ElseIf Me.IsArray AndAlso comparer.IsArray Then
                Return Me.ArrayMap?.Count < comparer.ArrayMap?.Count()
            Else
                Return Me.AsString() < comparer.AsString()
            End If
        End Function

        Public ReadOnly Property IsDate As Boolean
            Get
                Return numberType = NumberType.Date
            End Get
        End Property

        Public Function GreaterThan(comparer As Primitive) As Boolean
            If Me.IsNumber AndAlso comparer.IsNumber Then
                Return Me.AsDecimal() > comparer.AsDecimal()
            ElseIf Me.IsArray AndAlso comparer.IsArray Then
                Return Me.ArrayMap?.Count > comparer.ArrayMap?.Count()
            Else
                Return Me.AsString() > comparer.AsString()
            End If
        End Function

        Public Function LessThanOrEqualTo(comparer As Primitive) As Boolean
            If Me.IsNumber AndAlso comparer.IsNumber Then
                Return Me.AsDecimal() <= comparer.AsDecimal()
            ElseIf Me.IsArray AndAlso comparer.IsArray Then
                Return Me.ArrayMap?.Count <= comparer.ArrayMap?.Count()
            Else
                Return Me.AsString() <= comparer.AsString()
            End If
        End Function

        Public Function GreaterThanOrEqualTo(comparer As Primitive) As Boolean
            If Me.IsNumber AndAlso comparer.IsNumber Then
                Return Me.AsDecimal() >= comparer.AsDecimal()
            ElseIf Me.IsArray AndAlso comparer.IsArray Then
                Return Me.ArrayMap?.Count >= comparer.ArrayMap?.Count()
            Else
                Return Me.AsString() >= comparer.AsString()
            End If
        End Function

        Public Function EqualTo(comparer As Primitive) As Boolean
            Return AreEqual(Me, comparer)
        End Function

        Public Function NotEqualTo(comparer As Primitive) As Boolean
            Return Not AreEqual(Me, comparer)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If AsString() Is Nothing Then _stringValue = ""

            If TypeOf obj Is Primitive Then
                Dim p1 = CType(obj, Primitive)
                Dim s1 = p1.AsString()
                Dim s2 = Me.AsString()
                Dim b1 = s1.ToLower(CultureInfo.InvariantCulture)
                Dim b2 = s2.ToLower(CultureInfo.InvariantCulture)

                If s1 = "" Then
                    If s2 <> "" Then Return False
                ElseIf s2 = "" Then
                    Return False
                End If

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

        Friend Sub ConstructArrayMap()
            If _arrMap?.Count > 0 Then Return

            _arrMap = New Dictionary(Of Primitive, Primitive)(PrimitiveComparer.Instance)
            If IsEmpty Then Return

            Dim source = AsString().ToCharArray()
            Dim index As Integer = 0
            FillSubArray(_arrMap, source, index)
        End Sub

        Private Sub FillSubArray(
                           map As Dictionary(Of Primitive, Primitive),
                           source() As Char,
                           ByRef index As Integer
                     )

            Dim L = source.Length
            Do
                Dim key = Unescape(source, index)
                If key = "" Then Return

                If IsValidKey(key) Then
                    If index = L Then Return
                Else
                    map.Clear()
                    Return
                End If

                Dim value = Unescape(source, index, True)
                If value Is Nothing Then
                    map.Clear()
                    Return
                End If

                If source(index - 1) = "="c Then
                    index -= 2
                    Dim arr As New Primitive
                    Dim map2 As New Dictionary(Of Primitive, Primitive)(PrimitiveComparer.Instance)
                    FillSubArray(map2, source, index)
                    If map2.Count = 0 Then
                        map.Clear()
                        Return
                    End If
                    arr.ArrayMap = map2
                    map(New Primitive(key)) = arr
                Else
                    map(New Primitive(key)) = New Primitive(value)
                End If

            Loop While index < L

        End Sub

        Private Function IsValidKey(key As String) As Boolean
            For Each c In key
                Select Case c
                    Case " ", ".", "_"

                    Case Else
                        If Not (Char.IsDigit(c) OrElse Char.IsLetter(c)) Then
                            Return False
                        End If
                End Select
            Next

            Return True
        End Function

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
            Return System.Math.Round(primitive.AsDecimal(), MidpointRounding.AwayFromZero)
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
            If value.IsArray Then Return value.ArrayMap?.Count > 0
            Dim d = value.TryGetAsDecimal()
            Return d.HasValue  ' we tested for zoro before, so, if it is numeric it is non zero, if it's not numeric, it is any string which means false
        End Operator

        Public Shared Operator =(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return New Primitive(AreEqual(primitive1, primitive2))
        End Operator

        Private Shared Function AreEqual(p1 As Primitive, p2 As Primitive) As Boolean
            If p1.IsEmpty Then
                Return p2.IsEmpty
            ElseIf p2.IsEmpty Then
                Return False
            End If

            If p1.IsArray Then
                If Not p2.IsArray Then Return False
                Dim map1 = p1.ArrayMap
                Dim map2 = p2.ArrayMap
                Dim n = map1.Count
                If n <> map2.Count Then Return False

                Dim ids1 = map1.Keys
                Dim ids2 = map2.Keys

                For i = 0 To n - 1
                    Dim key = ids1(i)
                    If Not ids2.Contains(key) Then Return False
                    If Not AreEqual(map1(key), map2(key)) Then Return False
                Next
                Return True

            ElseIf Array.IsArray(p2) Then
                Return False
            Else
                Return p2.Equals(p1)
            End If
        End Function


        Public Shared Operator <>(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return New Primitive(Not AreEqual(primitive1, primitive2))
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

        Public Shared Operator &(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return New Primitive(primitive1.AsString() & primitive2.AsString())
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

        Public Shared Operator Mod(primitive1 As Primitive, primitive2 As Primitive) As Primitive
            Return primitive1.Remainder(primitive2)
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
                Return array.Items(indexer)
            Else
                Dim value As Primitive = Nothing
                If Not array.ArrayMap.TryGetValue(indexer, value) Then
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
            array.ArrayMap = New Dictionary(Of Primitive, Primitive)(PrimitiveComparer.Instance)

            For i = 0 To arr.Length - 1
                array.ArrayMap.Add(New Primitive((i + 1)), New Primitive(arr(i)))
            Next

            Return array
        End Function


        Public Shared Function SetArrayIndexer(
                        value As Primitive,
                        array As Primitive,
                        indexer As Primitive
                   ) As Primitive

            array.ConstructArrayMap()
            Dim map = array.ArrayMap
            Dim itemCount = If(map Is Nothing, 0, map.Count)

            If itemCount = 0 Then
                If CStr(array) <> "" Then
                    ' imdex the string
                    array.Items(indexer) = value
                    Return array
                Else
                    Dim map2 As New Dictionary(Of Primitive, Primitive)(PrimitiveComparer.Instance)
                    map2(indexer) = value
                    Return ConvertFromMap(map2)
                End If

            Else
                map(indexer) = value
                Return array
            End If
        End Function


        Public Shared Function SetArrayValue(
                        value As Primitive,
                        array As Primitive,
                        indexer As Primitive
                   ) As Primitive

            array.ConstructArrayMap()
            If array.GetItemCount = 0 AndAlso CStr(array) <> "" Then
                ' imdex the string
                array.Items(indexer) = value
                Return array
            Else
                Dim map As New Dictionary(Of Primitive, Primitive)(array.ArrayMap, PrimitiveComparer.Instance)
                map(indexer) = value
                Return ConvertFromMap(map)
            End If

        End Function

        Public Shared Function ConvertFromMap(map As Dictionary(Of Primitive, Primitive)) As Primitive
            Dim result As Primitive
            result.ArrayMap = map
            result._isArray = True
            Return result
        End Function

        Private Shared Function Escape(value As String) As String
            Dim stringBuilder1 As New StringBuilder
            For Each c As Char In value
                Select Case c
                    Case "="c
                        stringBuilder1.Append("\=")
                    Case ";"c
                        stringBuilder1.Append("\;")
                    Case "\"c
                        stringBuilder1.Append("\\")
                    Case Else
                        stringBuilder1.Append(c)
                End Select
            Next

            Return stringBuilder1.ToString()
        End Function

        Private Shared Function Unescape(source As Char(), ByRef index As Integer, Optional isValue As Boolean = False) As String
            Dim length = source.Length
            Dim sb As New StringBuilder

            While index < length
                Dim c As Char = source(index)
                index += 1
                If c = "\"c Then
                    If index < length Then
                        Dim c2 = source(index)
                        Select Case c2
                            Case "\"c, "="c, ";"c
                                sb.Append(c2)
                                index += 1
                                Continue While
                        End Select
                    End If

                    sb.Append(c)
                    Continue While
                End If

                If c = "="c Then
                    If sb.Length = 0 OrElse isValue Then Return Nothing
                    Return sb.ToString()
                ElseIf c = ";"c Then
                    Return sb.ToString()
                End If

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
