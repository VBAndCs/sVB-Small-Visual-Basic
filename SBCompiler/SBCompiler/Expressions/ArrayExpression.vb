Imports System
Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.Expressions
    <Serializable>
    Public Class ArrayExpression
        Inherits Expression

        Public Property LeftHand As Expression
        Public Property Indexer As Expression

        Public Shared Function ParseAndEmit(code As String, subroutine As SubroutineStatement, scope As CodeGenScope, lineOffset As Integer) As Expression
            Dim tempRoutine = SubroutineStatement.Current
            SubroutineStatement.Current = subroutine

            Dim _parser = Parser.Parse(code, scope.SymbolTable, lineOffset)

            Dim semantic As New SemanticAnalyzer(_parser, scope.TypeInfoBag)
            semantic.Analyze()

            'Build new fields
            For Each key In _parser.SymbolTable.GlobalVariables.Keys
                If Not scope.Fields.ContainsKey(key) Then
                    Dim fieldBuilder = scope.TypeBuilder.DefineField(key, GetType(Primitive), FieldAttributes.Private Or FieldAttributes.Static)
                    scope.Fields.Add(key, fieldBuilder)
                End If
            Next

            ' EmitIL
            For Each item In _parser.ParseTree
                item.PrepareForEmit(scope)
            Next

            For Each item In _parser.ParseTree
                item.EmitIL(scope)
            Next

            SubroutineStatement.Current = tempRoutine

            Return CType(_parser.ParseTree(0), AssignmentStatement).LeftValue
        End Function


        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            If LeftHand IsNot Nothing Then
                LeftHand.Parent = Me.Parent
                LeftHand.AddSymbols(symbolTable)
            End If

            If Indexer IsNot Nothing Then
                Indexer.Parent = Me.Parent
                Indexer.AddSymbols(symbolTable)
            End If
        End Sub

        Public Sub AddSymbolInitialization(symbolTable As SymbolTable)
            Dim arrayExpression = TryCast(LeftHand, ArrayExpression)
            Dim identifierExpression = TryCast(LeftHand, IdentifierExpression)

            If arrayExpression IsNot Nothing Then
                arrayExpression.AddSymbolInitialization(symbolTable)
            Else
                identifierExpression?.AddSymbolInitialization(symbolTable)
            End If
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If LeftHand Is Nothing Then Return

            LeftHand.EmitIL(scope)
            If Indexer Is Nothing Then
                ' This is just for intellisense info of the global file
                ' that may contain errors which is igmored at design time.
                ' This will not affect the compilation.
                LiteralExpression.Zero.EmitIL(scope)
            Else
                Indexer.EmitIL(scope)
            End If

            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.GetArrayValue, Nothing)
        End Sub

        Public Sub EmitILForSetter(scope As CodeGenScope)
            If LeftHand Is Nothing OrElse Indexer Is Nothing Then Return

            LeftHand.EmitIL(scope)
            Indexer.EmitIL(scope)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.SetArrayValue, Nothing)
            Dim arrayExpression = TryCast(LeftHand, ArrayExpression)
            Dim identifierExpression = TryCast(LeftHand, IdentifierExpression)

            If arrayExpression IsNot Nothing Then
                arrayExpression.EmitILForSetter(scope)
            ElseIf identifierExpression IsNot Nothing Then
                Dim subroutine = identifierExpression.Subroutine
                Dim identifier = identifierExpression.Identifier

                Dim var = scope.GetLocalBuilder(subroutine, identifier)
                If var IsNot Nothing Then
                    scope.ILGenerator.Emit(OpCodes.Stloc, var)

                ElseIf scope.Fields.ContainsKey(identifier.LCaseText) Then
                    Dim field = scope.Fields(identifierExpression.Identifier.LCaseText)
                    scope.ILGenerator.Emit(OpCodes.Stsfld, field)

                Else
                    scope.SymbolTable.Errors.Add(New [Error](identifier, $"The variable `{identifier.Text}` is used before being initialized."))
                End If

            End If
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0}[{1}]", New Object(1) {LeftHand.ToString(), Indexer})
        End Function
    End Class
End Namespace
