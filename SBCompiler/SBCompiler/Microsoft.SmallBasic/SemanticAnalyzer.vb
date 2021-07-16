Imports System
Imports System.Globalization
Imports System.Reflection
Imports Microsoft.SmallBasic.Expressions
Imports Microsoft.SmallBasic.Statements

Namespace Microsoft.SmallBasic
    Public Class SemanticAnalyzer
        Private _parser As Parser
        Friend _symbolTable As SymbolTable
        Friend _typeInfoBag As TypeInfoBag

        Public Sub New(parser As Parser, typeInfoBag As TypeInfoBag)
            If parser Is Nothing Then Throw New ArgumentNullException("parser")
            If typeInfoBag Is Nothing Then Throw New ArgumentNullException("typeInfoBag")

            _parser = parser
            _symbolTable = _parser.SymbolTable
            _typeInfoBag = typeInfoBag
        End Sub

        Public Sub Analyze()
            For Each item In _parser.ParseTree
                AnalyzeStatement(item)
            Next

            If _parser.Errors.Count <> 0 Then Return

            For Each variable In _symbolTable.Variables
                If Not _symbolTable.InitializedVariables.ContainsKey(variable.Key) Then
                    _parser.AddError(variable.Value, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("VariableNotInitialized"), New Object(0) {variable.Value.Text}))
                End If
            Next
        End Sub

        Private Sub AnalyzeExpression(ByVal expression As Expression, ByVal leaveValueInStack As Boolean, ByVal mustBeAssignable As Boolean)
            Dim type As Type = expression.GetType()

            If type Is GetType(BinaryExpression) Then
                AnalyzeBinaryExpression(TryCast(expression, BinaryExpression), leaveValueInStack, mustBeAssignable)
            ElseIf type Is GetType(ArrayExpression) Then
                AnalyzeArrayExpression(TryCast(expression, ArrayExpression), leaveValueInStack, mustBeAssignable)
            ElseIf type Is GetType(IdentifierExpression) Then
                AnalyzeIdentifierExpression(TryCast(expression, IdentifierExpression), leaveValueInStack, mustBeAssignable)
            ElseIf type Is GetType(MethodCallExpression) Then
                AnalyzeMethodCallExpression(TryCast(expression, MethodCallExpression), leaveValueInStack, mustBeAssignable)
            ElseIf type Is GetType(NegativeExpression) Then
                AnalyzeNegativeExpression(TryCast(expression, NegativeExpression), leaveValueInStack, mustBeAssignable)
            ElseIf type Is GetType(PropertyExpression) Then
                AnalyzePropertyExpression(TryCast(expression, PropertyExpression), leaveValueInStack, mustBeAssignable)
            End If
        End Sub

        Private Sub AnalyzeStatement(ByVal statement As Statement)
            Dim type As Type = statement.GetType()

            If type Is GetType(AssignmentStatement) Then
                AnalyzeAssignmentStatement(TryCast(statement, AssignmentStatement))
            ElseIf type Is GetType(ElseIfStatement) Then
                AnalyzeElseIfStatement(TryCast(statement, ElseIfStatement))
            ElseIf type Is GetType(ForStatement) Then
                AnalyzeForStatement(TryCast(statement, ForStatement))
            ElseIf type Is GetType(GotoStatement) Then
                AnalyzeGotoStatement(TryCast(statement, GotoStatement))
            ElseIf type Is GetType(IfStatement) Then
                AnalyzeIfStatement(TryCast(statement, IfStatement))
            ElseIf type Is GetType(MethodCallStatement) Then
                AnalyzeMethodCallStatement(TryCast(statement, MethodCallStatement))
            ElseIf type Is GetType(SubroutineCallStatement) Then
                AnalyzeSubroutineCallStatement(TryCast(statement, SubroutineCallStatement))
            ElseIf type Is GetType(SubroutineStatement) Then
                AnalyzeSubroutineStatement(TryCast(statement, SubroutineStatement))
            ElseIf type Is GetType(WhileStatement) Then
                AnalyzeWhileStatement(TryCast(statement, WhileStatement))
            End If
        End Sub

        Private Sub AnalyzeBinaryExpression(ByVal binaryExpression As BinaryExpression, ByVal leaveValueInStack As Boolean, ByVal mustBeAssignable As Boolean)
            If binaryExpression.LeftHandSide IsNot Nothing Then
                AnalyzeExpression(binaryExpression.LeftHandSide, leaveValueInStack, mustBeAssignable)
            End If

            If binaryExpression.RightHandSide IsNot Nothing Then
                AnalyzeExpression(binaryExpression.RightHandSide, leaveValueInStack, mustBeAssignable)
            End If
        End Sub

        Private Sub AnalyzeArrayExpression(ByVal arrayExpression As ArrayExpression, ByVal leaveValueInStack As Boolean, ByVal mustBeAssignable As Boolean)
            If arrayExpression.LeftHand IsNot Nothing Then
                AnalyzeExpression(arrayExpression.LeftHand, leaveValueInStack, mustBeAssignable)
            End If

            If arrayExpression.Indexer IsNot Nothing Then
                AnalyzeExpression(arrayExpression.Indexer, leaveValueInStack, mustBeAssignable)
            End If
        End Sub

        Private Sub AnalyzeIdentifierExpression(ByVal identifierExpression As IdentifierExpression, ByVal leaveValueInStack As Boolean, ByVal mustBeAssignable As Boolean)
            If identifierExpression.Identifier.Token <> 0 Then
                NoteVariableReference(identifierExpression.Identifier)
            End If
        End Sub

        Private Sub AnalyzeMethodCallExpression(ByVal methodCallExpression As MethodCallExpression, ByVal leaveValueInStack As Boolean, ByVal mustBeAssignable As Boolean)
            NoteMethodCallReference(methodCallExpression, leaveValueInStack, mustBeAssignable)

            For Each argument In methodCallExpression.Arguments
                AnalyzeExpression(argument, leaveValueInStack, mustBeAssignable)
            Next
        End Sub

        Private Sub AnalyzeNegativeExpression(ByVal negativeExpression As NegativeExpression, ByVal leaveValueInStack As Boolean, ByVal mustBeAssignable As Boolean)
            If negativeExpression.Expression IsNot Nothing Then
                AnalyzeExpression(negativeExpression.Expression, leaveValueInStack, mustBeAssignable)
            End If
        End Sub

        Private Sub AnalyzePropertyExpression(ByVal propertyExpression As PropertyExpression, ByVal leaveValueInStack As Boolean, ByVal mustBeAssignable As Boolean)
            NotePropertyReference(propertyExpression, leaveValueInStack, mustBeAssignable)
        End Sub

        Private Sub AnalyzeAssignmentStatement(ByVal assignmentStatement As AssignmentStatement)
            Dim identifierExpression = TryCast(assignmentStatement.RightValue, IdentifierExpression)

            If identifierExpression IsNot Nothing AndAlso _symbolTable.Subroutines.ContainsKey(identifierExpression.Identifier.NormalizedText) Then
                NoteEventReference(assignmentStatement.LeftValue, identifierExpression.Identifier)
                Return
            End If

            If assignmentStatement.LeftValue IsNot Nothing Then
                AnalyzeExpression(assignmentStatement.LeftValue, leaveValueInStack:=False, mustBeAssignable:=True)
            End If

            If assignmentStatement.RightValue IsNot Nothing Then
                AnalyzeExpression(assignmentStatement.RightValue, leaveValueInStack:=True, mustBeAssignable:=False)
            End If
        End Sub

        Private Sub AnalyzeElseIfStatement(ByVal elseIfStatement As ElseIfStatement)
            If elseIfStatement.Condition IsNot Nothing Then
                AnalyzeExpression(elseIfStatement.Condition, leaveValueInStack:=True, mustBeAssignable:=False)
            End If

            For Each thenStatement In elseIfStatement.ThenStatements
                AnalyzeStatement(thenStatement)
            Next
        End Sub

        Private Sub AnalyzeForStatement(ByVal forStatement As ForStatement)
            If forStatement.Iterator.Token <> 0 Then
                NoteVariableReference(forStatement.Iterator)
            End If

            If forStatement.InitialValue IsNot Nothing Then
                AnalyzeExpression(forStatement.InitialValue, leaveValueInStack:=True, mustBeAssignable:=False)
            End If

            If forStatement.FinalValue IsNot Nothing Then
                AnalyzeExpression(forStatement.FinalValue, leaveValueInStack:=True, mustBeAssignable:=False)
            End If

            If forStatement.StepValue IsNot Nothing Then
                AnalyzeExpression(forStatement.StepValue, leaveValueInStack:=True, mustBeAssignable:=False)
            End If

            For Each item In forStatement.ForBody
                AnalyzeStatement(item)
            Next
        End Sub

        Private Sub AnalyzeGotoStatement(ByVal gotoStatement As GotoStatement)
            If gotoStatement.Label.Token <> 0 Then
                NoteLabelReference(gotoStatement.Label)
            End If
        End Sub

        Private Sub AnalyzeIfStatement(ByVal ifStatement As IfStatement)
            If ifStatement.Condition IsNot Nothing Then
                AnalyzeExpression(ifStatement.Condition, leaveValueInStack:=True, mustBeAssignable:=False)
            End If

            For Each thenStatement In ifStatement.ThenStatements
                AnalyzeStatement(thenStatement)
            Next

            For Each elseIfStatement In ifStatement.ElseIfStatements
                AnalyzeStatement(elseIfStatement)
            Next

            For Each elseStatement In ifStatement.ElseStatements
                AnalyzeStatement(elseStatement)
            Next
        End Sub

        Private Sub AnalyzeMethodCallStatement(ByVal methodCallStatement As MethodCallStatement)
            If methodCallStatement.MethodCallExpression IsNot Nothing Then
                AnalyzeExpression(methodCallStatement.MethodCallExpression, leaveValueInStack:=False, mustBeAssignable:=False)
            End If
        End Sub

        Private Sub AnalyzeSubroutineCallStatement(ByVal subroutineCallStatement As SubroutineCallStatement)
            If subroutineCallStatement.SubroutineName.Token <> 0 Then
                NoteSubroutineCall(subroutineCallStatement.SubroutineName)
            End If
        End Sub

        Private Sub AnalyzeSubroutineStatement(ByVal subroutineStatement As SubroutineStatement)
            For Each item In subroutineStatement.SubroutineBody
                AnalyzeStatement(item)
            Next
        End Sub

        Private Sub AnalyzeWhileStatement(ByVal whileStatement As WhileStatement)
            If whileStatement.Condition IsNot Nothing Then
                AnalyzeExpression(whileStatement.Condition, leaveValueInStack:=True, mustBeAssignable:=False)
            End If

            For Each item In whileStatement.WhileBody
                AnalyzeStatement(item)
            Next
        End Sub

        Private Sub NoteEventReference(ByVal leftValue As Expression, ByVal subroutineName As TokenInfo)
            Dim propertyExpression As PropertyExpression = TryCast(leftValue, PropertyExpression)

            If propertyExpression IsNot Nothing Then
                Dim typeName = propertyExpression.TypeName
                Dim propertyName = propertyExpression.PropertyName

                If typeName.Token = Token.Illegal OrElse propertyName.Token = Token.Illegal Then
                    Return
                End If

                Dim value As TypeInfo = Nothing

                If _typeInfoBag.Types.TryGetValue(typeName.NormalizedText, value) Then
                    If Not value.Events.ContainsKey(propertyName.NormalizedText) Then
                        _parser.AddError(propertyName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("EventNotFound"), New Object(1) {propertyName.Text, typeName.Text}))
                    End If
                Else
                    _parser.AddError(typeName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("TypeNotFound"), New Object(0) {typeName.Text}))
                End If
            Else
                _parser.AddError(subroutineName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("SubroutineEventAssignment"), New Object(0) {subroutineName.Text}))
            End If
        End Sub

        Private Sub NoteLabelReference(ByVal label As TokenInfo)
            If label.Token <> 0 AndAlso Not _symbolTable.Labels.ContainsKey(label.NormalizedText) Then
                _parser.AddError(label, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("LabelNotFound"), New Object(0) {label.Text}))
            End If
        End Sub

        Private Sub NoteMethodCallReference(ByVal methodExpression As MethodCallExpression, ByVal leaveValueInStack As Boolean, ByVal isAssignable As Boolean)
            Dim typeName = methodExpression.TypeName
            Dim methodName = methodExpression.MethodName

            If typeName.Token = Token.Illegal OrElse methodName.Token = Token.Illegal Then
                Return
            End If

            Dim value As TypeInfo = Nothing

            If _typeInfoBag.Types.TryGetValue(typeName.NormalizedText, value) Then
                Dim value2 As MethodInfo = Nothing

                If value.Methods.TryGetValue(methodName.NormalizedText, value2) Then
                    Dim num As Integer = value2.GetParameters().Length

                    If num <> methodExpression.Arguments.Count Then
                        _parser.AddError(methodName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("ArgumentNumberMismatch"), typeName.Text, methodName.Text, methodExpression.Arguments.Count, num))
                    End If

                                        ''' Cannot convert IfStatementSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax' to type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax'.
'''    at ICSharpCode.CodeConverter.VB.MethodBodyExecutableStatementVisitor.VisitIfStatement(IfStatementSyntax node)
'''    at Microsoft.CodeAnalysis.CSharp.Syntax.IfStatementSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)
''' 
''' Input:
''' 					if (leaveValueInStack && value2.ReturnType == typeof(void))
''' 					{
''' 						this._parser.AddError(methodName, string.Format(System.Globalization.CultureInfo.CurrentUICulture, Microsoft.SmallBasic.ResourceHelper.GetString("ReturnValueExpectedFromVoidMethod"), new object[2]
''' 						{
''' 							typeName.Text,
''' 							methodName.Text
''' 						}));
''' 					}
''' 
''' 
                Else
                    _parser.AddError(methodName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("MethodNotFound"), New Object(1) {methodName.Text, typeName.Text}))
                End If
            Else
                _parser.AddError(typeName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("TypeNotFound"), New Object(0) {typeName.Text}))
            End If
        End Sub

        Private Sub NotePropertyReference(ByVal propertyExpression As PropertyExpression, ByVal leaveValueInStack As Boolean, ByVal mustBeAssignable As Boolean)
            Dim typeName = propertyExpression.TypeName
            Dim propertyName = propertyExpression.PropertyName

            If typeName.Token = Token.Illegal OrElse propertyName.Token = Token.Illegal Then
                Return
            End If

            Dim value As TypeInfo = Nothing

            If _typeInfoBag.Types.TryGetValue(typeName.NormalizedText, value) Then
                Dim value2 As PropertyInfo = Nothing

                If value.Properties.TryGetValue(propertyName.NormalizedText, value2) Then
                    If mustBeAssignable AndAlso Not value2.CanWrite Then
                        _parser.AddError(propertyName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("PropertyIsReadOnly"), New Object(1) {propertyName.Text, typeName.Text}))
                    End If

                    If leaveValueInStack AndAlso Not value2.CanRead Then
                        _parser.AddError(propertyName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("PropertyIsWriteOnly"), New Object(1) {propertyName.Text, typeName.Text}))
                    End If
                Else
                    _parser.AddError(propertyName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("PropertyNotFound"), New Object(1) {propertyName.Text, typeName.Text}))
                End If
            Else
                _parser.AddError(typeName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("TypeNotFound"), New Object(0) {typeName.Text}))
            End If
        End Sub

        Private Sub NoteSubroutineCall(ByVal subroutineName As TokenInfo)
            If subroutineName.Token <> 0 AndAlso Not _symbolTable.Subroutines.ContainsKey(subroutineName.NormalizedText) Then
                _parser.AddError(subroutineName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("SubroutineNotDefined"), New Object(0) {subroutineName.Text}))
            End If
        End Sub

        Private Sub NoteVariableReference(ByVal variable As TokenInfo)
            If variable.Token <> 0 AndAlso Not _symbolTable.Variables.ContainsKey(variable.NormalizedText) AndAlso _symbolTable.Subroutines.ContainsKey(variable.NormalizedText) Then
                _parser.AddError(variable, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("SubroutineUsedAsVariable"), New Object(0) {variable.Text}))
            End If
        End Sub
    End Class
End Namespace
