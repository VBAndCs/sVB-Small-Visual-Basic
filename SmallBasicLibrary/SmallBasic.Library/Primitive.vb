
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
                Return String.Equals(x.AsString, y.AsString, StringComparison.InvariantCultureIgnoreCase)
            End Function

            Public Function GetHashCode(obj As Primitive) As Integer Implements Collections.Generic.IEqualityComparer(Of Microsoft.SmallBasic.Library.Primitive).GetHashCode
                Return obj.GetHashCode()
            End Function
        End Class


        Private _stringValue As String

        Private _decimalValue As Decimal?

        Friend _arrayMap As Dictionary(Of Primitive, Primitive)
        Friend _isArray As Boolean

        Default Public Property Items(index As Primitive) As Primitive
            Get
                ConstructArrayMap()
                If _arrayMap.Count = 0 Then
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
                If _arrayMap.Count = 0 AndAlso _stringValue <> "" Then
                    If index < 1 Then
                        Me._stringValue = CStr(Value) & Space(-index) & _stringValue
                    ElseIf index > _stringValue.Length Then
                        Me._stringValue &= Space(index - _stringValue.Length - 1) & CStr(Value)
                    Else
                        Me._stringValue = _stringValue.Substring(0, index - 1) & CStr(Value) & _stringValue.Substring(index)
                    End If

                Else
                    Dim primitive As Primitive = SetArrayValue(Value, Me, index)
                    _stringValue = primitive._stringValue
                    _arrayMap = primitive._arrayMap
                    _decimalValue = primitive._decimalValue
                End If
            End Set
        End Property

        Friend ReadOnly Property AsString As String
            Get
                If _stringValue IsNot Nothing Then Return _stringValue

                If _decimalValue.HasValue Then
                    _stringValue = _decimalValue.Value.ToString()

                ElseIf _arrayMap IsNot Nothing Then
                    Dim sb As New StringBuilder
                    For Each entry In _arrayMap
                        sb.AppendFormat("{0}={1};", Escape(entry.Key), Escape(entry.Value))
                    Next

                    _stringValue = sb.ToString()
                End If

                Return _stringValue
            End Get
        End Property

        Friend ReadOnly Property IsArray As Boolean
            Get
                If _isArray Then Return True
                ConstructArrayMap()
                Return _arrayMap.Count > 0
            End Get
        End Property

        Friend ReadOnly Property IsEmpty As Boolean
            Get
                If String.IsNullOrEmpty(_stringValue) AndAlso Not _decimalValue.HasValue Then
                    If _arrayMap IsNot Nothing Then
                        Return _arrayMap.Count = 0
                    End If

                    Return True
                End If

                Return False
            End Get
        End Property

        Friend ReadOnly Property IsNumber As Boolean
            Get
                Dim result As Decimal = 0D
                If Not _decimalValue.HasValue Then
                    Return Decimal.TryParse(AsString, NumberStyles.Float, CultureInfo.InvariantCulture, result)
                End If

                Return True
            End Get
        End Property

        Public Sub New(primitiveText As String)
            If primitiveText Is Nothing Then
                Throw New ArgumentNullException("primitiveText")
            End If
            _stringValue = primitiveText
            _decimalValue = Nothing
            _arrayMap = Nothing
        End Sub

        Public Sub New(primitiveInteger As Integer)
            _decimalValue = primitiveInteger
            _stringValue = Nothing
            _arrayMap = Nothing
        End Sub

        Public Sub New(primitiveDecimal As Decimal)
            _decimalValue = primitiveDecimal
            _stringValue = Nothing
            _arrayMap = Nothing
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
            If CBool(primitiveBool) Then
                _stringValue = "True"
            Else
                _stringValue = "False"
            End If
            _arrayMap = Nothing
        End Sub

        Public Function Append(primitive As Primitive) As Primitive
            Return New Primitive(AsString & primitive.AsString)
        End Function

        Public Function Add(addend As Primitive) As Primitive
            Dim n1 = TryGetAsDecimal()
            Dim n2 = addend.TryGetAsDecimal()
            If n1.HasValue AndAlso n2.HasValue Then
                Return n1.Value + n2.Value
            End If
            Return AsString() & addend.AsString()
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
            Dim dictionary1 As New Dictionary(Of Primitive, Primitive)(_arrayMap.Count, PrimitiveComparer.Instance)
            Dim num As Integer = 1
            For Each key1 As Primitive In _arrayMap.Keys
                dictionary1(num) = key1
                num += 1
            Next

            Return ConvertFromMap(dictionary1)
        End Function

        Public Function GetItemCount() As Primitive
            ConstructArrayMap()
            Return _arrayMap.Count
        End Function

        Public Function Subtract(addend As Primitive) As Primitive
            Return New Primitive(GetAsDecimal() - addend.GetAsDecimal())
        End Function

        Public Function Multiply(multiplicand As Primitive) As Primitive
            Return New Primitive(GetAsDecimal() * multiplicand.GetAsDecimal())
        End Function

        Public Function Divide(divisor As Primitive) As Primitive
            divisor.GetAsDecimal()
            Return New Primitive(GetAsDecimal() / divisor.GetAsDecimal())
        End Function

        Public Function LessThan(comparer As Primitive) As Boolean
            Return GetAsDecimal() < comparer.GetAsDecimal()
        End Function

        Public Function GreaterThan(comparer As Primitive) As Boolean
            Return GetAsDecimal() > comparer.GetAsDecimal()
        End Function

        Public Function LessThanOrEqualTo(comparer As Primitive) As Boolean
            Return GetAsDecimal() <= comparer.GetAsDecimal()
        End Function

        Public Function GreaterThanOrEqualTo(comparer As Primitive) As Boolean
            Return GetAsDecimal() >= comparer.GetAsDecimal()
        End Function

        Public Function EqualTo(comparer As Primitive) As Boolean
            Return Equals(comparer)
        End Function

        Public Function NotEqualTo(comparer As Primitive) As Boolean
            Return Not Equals(comparer)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If AsString Is Nothing Then
                _stringValue = ""
            End If

            If TypeOf obj Is Primitive Then
                Dim p1 = CType(obj, Primitive)
                Dim s1 = p1.AsString()
                Dim s2 = Me.AsString()
                Dim b1 = s1.ToLower(CultureInfo.InvariantCulture)
                Dim b2 = s2.ToLower(CultureInfo.InvariantCulture)

                If s1 = "" Then
                    If s2 = "" OrElse b1 = "false" Then Return True
                ElseIf b1 = "true" Then
                    Return If(Me.IsNumber, Me.GetAsDecimal() <> 0, b2 = "true")
                ElseIf b1 = "false" Then
                    Return If(Me.IsNumber, Me.GetAsDecimal() = 0, b2 = "" OrElse b2 = "false")
                ElseIf Me.IsNumber AndAlso p1.IsNumber Then
                    Return Me.GetAsDecimal() = p1.GetAsDecimal()
                Else
                    Return s1 = s2
                End If
            End If

            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            If AsString Is Nothing Then
                _stringValue = ""
            End If

            Return AsString.ToUpper(CultureInfo.InvariantCulture).GetHashCode()
        End Function

        Public Overrides Function ToString() As String
            Return AsString
        End Function

        Private Sub ConstructArrayMap()
            If _arrayMap IsNot Nothing Then Return
            _arrayMap = New Dictionary(Of Primitive, Primitive)(PrimitiveComparer.Instance)

            If IsEmpty Then Return

            Dim source = AsString.ToCharArray()

            Dim index As Integer = 0
            Do
                Dim key = Unescape(source, index)
                If key = "" Then Exit Do

                Dim value = Unescape(source, index)
                If value = "" Then Exit Do

                _arrayMap(key) = value
            Loop
        End Sub

        Friend Function TryGetAsDecimal() As Decimal?
            If IsEmpty Then
                Return Nothing
            End If
            If _decimalValue.HasValue Then
                Return _decimalValue
            End If
            Dim result As Decimal = 0D
            If Decimal.TryParse(AsString, NumberStyles.Float, CultureInfo.InvariantCulture, result) Then
                _decimalValue = result
            End If

            Return _decimalValue
        End Function

        Friend Function GetAsDecimal() As Decimal
            If IsEmpty Then
                Return 0D
            End If
            If _decimalValue.HasValue Then
                Return _decimalValue.Value
            End If
            Dim result As Decimal = 0D
            If Decimal.TryParse(AsString, NumberStyles.Float, CultureInfo.InvariantCulture, result) Then
                _decimalValue = result
            End If

            Return result
        End Function

        Public Shared Function ConvertToBoolean(primitive As Primitive) As Boolean
            Return primitive
        End Function

        Public Shared Widening Operator CType(primitive1 As Primitive) As String
            If primitive1.AsString Is Nothing Then
                Return ""
            End If

            Return primitive1.AsString
        End Operator

        Public Shared Widening Operator CType(primitive1 As Primitive) As Integer
            Return CInt(primitive1.GetAsDecimal())
        End Operator

        Public Shared Widening Operator CType(primitive1 As Primitive) As Single
            Return CSng(primitive1.GetAsDecimal())
        End Operator

        Public Shared Widening Operator CType(primitive1 As Primitive) As Double
            Return CDbl(primitive1.GetAsDecimal())
        End Operator

        Public Shared Widening Operator CType(value As Primitive) As Boolean
            Dim s = CStr(value)
            If s Is Nothing Then Return False
            If value.IsArray AndAlso value._arrayMap?.Count > 0 Then Return True
            If s.ToLower() = "true" Then Return True
            If Not IsNumeric(s) Then Return False
            If Double.Parse(s) = 0 Then Return False
            Return True
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
            Return primitive1.LessThanOrEqualTo(primitive2)
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
            Return -primitive1.GetAsDecimal()
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

            Return New Primitive(value)
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
                Dim dictionary1 As New Dictionary(Of Primitive, Primitive)(array._arrayMap, PrimitiveComparer.Instance)
                If value.IsEmpty Then
                    dictionary1.Remove(indexer)
                Else
                    dictionary1(indexer) = value
                End If

                Return ConvertFromMap(dictionary1)
            End If

        End Function

        Public Shared Function ConvertFromMap(map As Dictionary(Of Primitive, Primitive)) As Primitive
            Dim result As Primitive = CType(Nothing, Primitive)
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
            Dim flag As Boolean = False
            Dim flag2 As Boolean = True
            Dim num As Integer = source.Length
            Dim stringBuilder1 As New StringBuilder
            While index < num
                Dim c As Char = source(index)
                index += 1
                If Not flag Then
                    If c = "\"c Then
                        flag = True
                        Continue While
                    End If
                    If c = "="c OrElse c = ";"c Then
                        Exit While
                    End If
                Else
                    flag = False
                End If
                flag2 = False
                stringBuilder1.Append(c)
            End While
            If flag2 Then
                Return Nothing
            End If

            Return stringBuilder1.ToString()
        End Function
    End Structure
End Namespace
