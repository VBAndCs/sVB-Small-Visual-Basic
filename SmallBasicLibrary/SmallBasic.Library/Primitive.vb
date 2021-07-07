
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


        Private _primitive As String

        Private _primitiveAsDecimal As Decimal?

        Private _arrayMap As Dictionary(Of Primitive, Primitive)

        Default Public Property item(index As Primitive) As Primitive
            Get
                If CBool(ContainsKey(index)) Then
                    Return _arrayMap(index)
                End If

                Return ""
            End Get
            Set(Value As Primitive)
                Dim primitive1 As Primitive = SetArrayValue(Value, Me, index)
                _primitive = primitive1._primitive
                _arrayMap = primitive1._arrayMap
                _primitiveAsDecimal = primitive1._primitiveAsDecimal
            End Set
        End Property
        Friend ReadOnly Property AsString As String
            Get
                If _primitive IsNot Nothing Then
                    Return _primitive
                End If

                If _primitiveAsDecimal.HasValue Then
                    _primitive = _primitiveAsDecimal.Value.ToString()
                End If

                If _arrayMap IsNot Nothing Then
                    Dim stringBuilder1 As New StringBuilder
                    For Each item As KeyValuePair(Of Primitive, Primitive) In _arrayMap
                        stringBuilder1.AppendFormat("{0}={1};", Escape(item.Key), Escape(item.Value))
                    Next

                    _primitive = stringBuilder1.ToString()
                End If

                Return _primitive
            End Get
        End Property
        Friend ReadOnly Property IsArray As Boolean
            Get
                ConstructArrayMap()
                Return _arrayMap.Count > 0
            End Get
        End Property
        Friend ReadOnly Property IsEmpty As Boolean
            Get
                If String.IsNullOrEmpty(_primitive) AndAlso Not _primitiveAsDecimal.HasValue Then
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
                If Not _primitiveAsDecimal.HasValue Then
                    Return Decimal.TryParse(AsString, NumberStyles.Float, CultureInfo.InvariantCulture, result)
                End If

                Return True
            End Get
        End Property

        Public Sub New(primitiveText As String)
            If primitiveText Is Nothing Then
                Throw New ArgumentNullException("primitiveText")
            End If
            _primitive = primitiveText
            _primitiveAsDecimal = Nothing
            _arrayMap = Nothing
        End Sub

        Public Sub New(primitiveInteger As Integer)
            _primitiveAsDecimal = primitiveInteger
            _primitive = Nothing
            _arrayMap = Nothing
        End Sub

        Public Sub New(primitiveDecimal As Decimal)
            _primitiveAsDecimal = primitiveDecimal
            _primitive = Nothing
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
            _primitiveAsDecimal = primitiveLong
            _primitive = Nothing
            _arrayMap = Nothing
        End Sub

        Public Sub New(primitiveObject As Object)
            _primitive = primitiveObject.ToString()
            _primitiveAsDecimal = Nothing
            _arrayMap = Nothing
        End Sub

        Public Sub New(primitiveBool As Boolean)
            _primitiveAsDecimal = Nothing
            If primitiveBool Then
                _primitive = "True"
            Else
                _primitive = "False"
            End If
            _arrayMap = Nothing
        End Sub

        Public Function Append(primitive1 As Primitive) As Primitive
            Return New Primitive(AsString & primitive1.AsString)
        End Function

        Public Function Add(addend As Primitive) As Primitive
            Dim num As Decimal? = TryGetAsDecimal()
            Dim num2 As Decimal? = addend.TryGetAsDecimal()
            If num.HasValue AndAlso num2.HasValue Then
                Return num.Value + num2.Value
            End If
            Return AsString & addend.AsString
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
                _primitive = ""
            End If
            If TypeOf obj Is Primitive Then
                Dim primitive1 As Primitive = CType(obj, Primitive)
                If primitive1.AsString IsNot Nothing AndAlso primitive1.AsString.ToLower(CultureInfo.InvariantCulture) = "true" AndAlso AsString IsNot Nothing AndAlso AsString.ToLower(CultureInfo.InvariantCulture) = "true" Then
                    Return True
                End If
                If IsNumber AndAlso primitive1.IsNumber Then
                    Return GetAsDecimal() = primitive1.GetAsDecimal()
                End If

                Return AsString = primitive1.AsString
            End If

            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            If AsString Is Nothing Then
                _primitive = ""
            End If

            Return AsString.ToUpper(CultureInfo.InvariantCulture).GetHashCode()
        End Function

        Public Overrides Function ToString() As String
            Return AsString
        End Function

        Private Sub ConstructArrayMap()
            If _arrayMap IsNot Nothing Then
                Return
            End If
            _arrayMap = New Dictionary(Of Primitive, Primitive)(PrimitiveComparer.Instance)
            If IsEmpty Then
                Return
            End If
            Dim source As Char() = AsString.ToCharArray()
            Dim index As Integer = 0
            While True
                Dim value As String = Unescape(source, index)
                If String.IsNullOrEmpty(value) Then
                    Exit While
                End If
                Dim text As String = Unescape(source, index)
                If text Is Nothing Then
                    Exit While
                End If
                _arrayMap(value) = text
            End While
        End Sub

        Friend Function TryGetAsDecimal() As Decimal?
            If IsEmpty Then
                Return Nothing
            End If
            If _primitiveAsDecimal.HasValue Then
                Return _primitiveAsDecimal
            End If
            Dim result As Decimal = 0D
            If Decimal.TryParse(AsString, NumberStyles.Float, CultureInfo.InvariantCulture, result) Then
                _primitiveAsDecimal = result
            End If

            Return _primitiveAsDecimal
        End Function

        Friend Function GetAsDecimal() As Decimal
            If IsEmpty Then
                Return 0D
            End If
            If _primitiveAsDecimal.HasValue Then
                Return _primitiveAsDecimal.Value
            End If
            Dim result As Decimal = 0D
            If Decimal.TryParse(AsString, NumberStyles.Float, CultureInfo.InvariantCulture, result) Then
                _primitiveAsDecimal = result
            End If

            Return result
        End Function

        Public Shared Function ConvertToBoolean(primitive1 As Primitive) As Boolean
            Return primitive1
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

        Public Shared Widening Operator CType(primitive1 As Primitive) As Boolean
            If primitive1.AsString IsNot Nothing AndAlso primitive1.AsString.Equals("true", StringComparison.InvariantCultureIgnoreCase) Then
                Return True
            End If

            Return False
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

        Public Shared Operator IsTrue(primitive1 As Primitive) As Boolean
            Return primitive1
        End Operator

        Public Shared Operator IsFalse(primitive1 As Primitive) As Boolean
            Return Not CBool(primitive1)
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
            Dim value As Microsoft.SmallBasic.Library.Primitive = Nothing
            If Not array._arrayMap.TryGetValue(indexer, value) Then
                value = CType(Nothing, Primitive)
            End If

            Return value
        End Function

        Public Shared Function SetArrayValue(value As Primitive, array As Primitive, indexer As Primitive) As Primitive
            array.ConstructArrayMap()
            Dim dictionary1 As New Dictionary(Of Primitive, Primitive)(array._arrayMap, PrimitiveComparer.Instance)
            If value.IsEmpty Then
                dictionary1.Remove(indexer)
            Else
                dictionary1(indexer) = value
            End If

            Return ConvertFromMap(dictionary1)
        End Function

        Public Shared Function ConvertFromMap(map As Dictionary(Of Primitive, Primitive)) As Primitive
            Dim result As Primitive = CType(Nothing, Primitive)
            result._primitive = Nothing
            result._primitiveAsDecimal = Nothing
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
