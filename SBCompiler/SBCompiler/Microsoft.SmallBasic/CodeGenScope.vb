Imports System.Collections.Generic
Imports System.Reflection
Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic
    Public Class CodeGenScope
        Private _fields As Dictionary(Of String, FieldInfo)
        Private _methodBuilders As Dictionary(Of String, MethodBuilder)
        Private _labels As Dictionary(Of String, Label)
        Private _parent As CodeGenScope
        Public Property TypeBuilder As TypeBuilder
        Public Property MethodBuilder As MethodBuilder
        Public Property ILGenerator As ILGenerator

        Public Property Parent As CodeGenScope
            Get
                Return _parent
            End Get
            Set(ByVal value As CodeGenScope)
                _parent = value
                _fields = value.Fields
                _labels = value.Labels
                _methodBuilders = value.MethodBuilders
                SymbolTable = value.SymbolTable
                TypeInfoBag = value.TypeInfoBag
            End Set
        End Property

        Public Property SymbolTable As SymbolTable
        Public Property TypeInfoBag As TypeInfoBag

        Public ReadOnly Property Fields As Dictionary(Of String, FieldInfo)
            Get

                If _fields Is Nothing Then
                    _fields = New Dictionary(Of String, FieldInfo)()
                End If

                Return _fields
            End Get
        End Property

        Public ReadOnly Property MethodBuilders As Dictionary(Of String, MethodBuilder)
            Get

                If _methodBuilders Is Nothing Then
                    _methodBuilders = New Dictionary(Of String, MethodBuilder)()
                End If

                Return _methodBuilders
            End Get
        End Property

        Public ReadOnly Property Labels As Dictionary(Of String, Label)
            Get

                If _labels Is Nothing Then
                    _labels = New Dictionary(Of String, Label)()
                End If

                Return _labels
            End Get
        End Property
    End Class
End Namespace
