Imports Microsoft.Contracts

Namespace Microsoft.Nautilus.Core
    Public NotInheritable Class TypeHelper

        Public Shared Function IsIdentifiedType(identifiedType As String, type1 As Type) As Boolean
            Contract.RequiresNotNull(identifiedType, "identifiedType")
            Contract.RequiresNotNull(type1, "type")
            Return String.CompareOrdinal(identifiedType, type1.AssemblyQualifiedName) = 0
        End Function

        Public Shared Function IsIdentifiedTypeAssignableFrom(identifiedType As String, type1 As Type) As Boolean
            Contract.RequiresNotNull(identifiedType, "identifiedType")
            Contract.RequiresNotNull(type1, "type")
            If Not IsTypeInheritedFromIdentifiedType(identifiedType, type1) Then
                Return IsTypeAnImplementationOfIdentifiedType(identifiedType, type1)
            End If
            Return True
        End Function

        Private Shared Function IsTypeInheritedFromIdentifiedType(identifiedType As String, type1 As Type) As Boolean
            Dim type2 As Type = type1
            While type2 IsNot Nothing
                If IsIdentifiedType(identifiedType, type2) Then
                    Return True
                End If
                type2 = type2.BaseType
            End While
            Return False
        End Function

        Private Shared Function IsTypeAnImplementationOfIdentifiedType(identifiedType As String, type1 As Type) As Boolean
            Dim interfaces As Type() = type1.GetInterfaces()
            For Each type2 As Type In interfaces
                If IsIdentifiedType(identifiedType, type2) Then
                    Return True
                End If
            Next
            Return False
        End Function
    End Class
End Namespace
