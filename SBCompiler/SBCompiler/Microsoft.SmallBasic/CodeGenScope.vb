Imports System.Collections.Generic
Imports System.Reflection
Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic
    Public Class CodeGenScope
        Private _fields As Dictionary(Of String, FieldInfo)
        Private Locals As New Dictionary(Of String, LocalBuilder)
        Private _methodBuilders As Dictionary(Of String, MethodBuilder)
        Private _labels As Dictionary(Of String, Label)
        Private _parent As CodeGenScope
        Public Property TypeBuilder As TypeBuilder
        Public Property MethodBuilder As MethodBuilder
        Public Property ILGenerator As ILGenerator

        Friend Function CreateLocalBuilder(Subroutine As Statements.SubroutineStatement, varIdentifier As TokenInfo, varType As Type) As LocalBuilder
            Dim varBuilder = GetLocalBuilder(Subroutine, varIdentifier)
            If varBuilder Is Nothing Then
                varBuilder = ILGenerator.DeclareLocal(varType)
                varBuilder.SetLocalSymInfo(varIdentifier.NormalizedText)
                AddLocalBuilder(Subroutine, varIdentifier, varBuilder)
            End If
            Return varBuilder
        End Function

        Friend Function GetLocalBuilder(Subroutine As Statements.SubroutineStatement, varIdentifier As TokenInfo) As LocalBuilder
            Dim key = ""
            If Subroutine Is Nothing Then
                key = varIdentifier.NormalizedText
            Else
                key = $"{Subroutine.Name.NormalizedText}.{varIdentifier.NormalizedText}"
            End If

            If Not Locals.ContainsKey(key) Then Return Nothing
            Return Locals(key)
        End Function

        Private Sub AddLocalBuilder(Subroutine As Statements.SubroutineStatement, varIdentifier As TokenInfo, localVarBuilder As LocalBuilder)
            Dim key = ""
            If Subroutine Is Nothing Then
                key = varIdentifier.NormalizedText
            Else
                key = $"{Subroutine.Name.NormalizedText}.{varIdentifier.NormalizedText}"
            End If

            Locals(key) = localVarBuilder
        End Sub


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
