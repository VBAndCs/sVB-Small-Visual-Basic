Imports System.Collections.Generic
Imports System.Reflection
Imports System.Reflection.Emit

Namespace Microsoft.SmallVisualBasic
    Public Class CodeGenScope
        Private _locals As New Dictionary(Of String, LocalBuilder)

        Private _parent As CodeGenScope
        Public Property TypeBuilder As TypeBuilder
        Public Property MethodBuilder As MethodBuilder
        Public Property ILGenerator As ILGenerator
        Public Property SymbolTable As SymbolTable
        Public Property TypeInfoBag As TypeInfoBag
        Public ReadOnly Property Fields As New Dictionary(Of String, FieldInfo)

        Public ReadOnly Property MethodBuilders As New Dictionary(Of String, MethodBuilder)

        Public ReadOnly Property Labels As New Dictionary(Of String, Label)

        Friend Function CreateLocalBuilder(
                          Subroutine As Statements.SubroutineStatement,
                          varIdentifier As Token
                    ) As LocalBuilder

            Dim varBuilder = GetLocalBuilder(Subroutine, varIdentifier)
            If varBuilder Is Nothing Then
                varBuilder = _ILGenerator.DeclareLocal(GetType(Library.Primitive))
                varBuilder.SetLocalSymInfo(varIdentifier.LCaseText)
                AddLocalBuilder(Subroutine, varIdentifier, varBuilder)
            End If
            Return varBuilder
        End Function


        Friend Function GetLocalBuilder(Subroutine As Statements.SubroutineStatement, varIdentifier As Token) As LocalBuilder
            Dim key = ""
            If Subroutine Is Nothing Then
                key = varIdentifier.LCaseText
            Else
                key = $"{Subroutine.Name.LCaseText}.{varIdentifier.LCaseText}"
            End If

            If Not _SymbolTable.LocalVariables.ContainsKey(key) Then Return Nothing

            If _locals.ContainsKey(key) Then Return _locals(key)

            Dim varBuilder = _ILGenerator.DeclareLocal(GetType(Library.Primitive))
            Dim var = _SymbolTable.LocalVariables(key)
            varBuilder.SetLocalSymInfo(var.Identifier.LCaseText)
            _locals(key) = varBuilder
            Return varBuilder

        End Function

        Private Sub AddLocalBuilder(
                           Subroutine As Statements.SubroutineStatement,
                           varIdentifier As Token,
                           localVarBuilder As LocalBuilder
                    )

            Dim key = ""
            If Subroutine Is Nothing Then
                key = varIdentifier.LCaseText
            Else
                key = $"{Subroutine.Name.LCaseText}.{varIdentifier.LCaseText}"
            End If

            _locals(key) = localVarBuilder
        End Sub


        Public Property Parent As CodeGenScope
            Get
                Return _parent
            End Get

            Set(value As CodeGenScope)
                _parent = value
                _fields = value._Fields
                _Labels = value._Labels
                _locals = value._locals
                _MethodBuilders = value.MethodBuilders
                SymbolTable = value.SymbolTable
                TypeInfoBag = value.TypeInfoBag
                _GlobalScope = value.GlobalScope
                _ForGlobalHelp = value.ForGlobalHelp
            End Set
        End Property

        Public Property GlobalScope As CodeGenScope
        Public Property ForGlobalHelp As Boolean
    End Class
End Namespace
