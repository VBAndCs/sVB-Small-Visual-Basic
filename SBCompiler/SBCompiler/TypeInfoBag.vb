Imports System
Imports System.Collections.Generic
Imports System.Reflection

Namespace Microsoft.SmallVisualBasic
    <Serializable>
    Public Class TypeInfoBag
        Inherits MarshalByRefObject

        Private _types As New Dictionary(Of String, TypeInfo)()
        Public StringToPrimitive As MethodInfo
        Public NumberToPrimitive As MethodInfo
        Public DateToPrimitive As MethodInfo
        Public TimeSpanToPrimitive As MethodInfo
        Public PrimitiveToBoolean As MethodInfo
        Public Negation As MethodInfo
        Public Add As MethodInfo
        Public Subtract As MethodInfo
        Public Multiply As MethodInfo
        Public Divide As MethodInfo
        Public EqualTo As MethodInfo
        Public NotEqualTo As MethodInfo
        Public LessThan As MethodInfo
        Public LessThanOrEqualTo As MethodInfo
        Public GreaterThan As MethodInfo
        Public GreaterThanOrEqualTo As MethodInfo
        Public [And] As MethodInfo
        Public [Or] As MethodInfo
        Public GetArrayValue As MethodInfo
        Public SetArrayValue As MethodInfo

        Public ReadOnly Property Types As Dictionary(Of String, TypeInfo)
            Get
                Return _types
            End Get
        End Property
    End Class
End Namespace
