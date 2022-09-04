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

            If _parser.Errors.Count > 0 Then Return

            For Each variable In _symbolTable.GlobalVariables
                If Not _symbolTable.InitializedVariables.ContainsKey(variable.Key) Then
                    _parser.AddError(variable.Value, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("VariableNotInitialized"), New Object(0) {variable.Value.Text}))
                End If
            Next
        End Sub

        Private Sub AnalyzeExpression(expression As Expression, leaveValueInStack As Boolean, mustBeAssignable As Boolean)
            Dim type As Type = expression.GetType()

            If type Is GetType(BinaryExpression) Then
                AnalyzeBinaryExpression(CType(expression, BinaryExpression), leaveValueInStack, mustBeAssignable)

            ElseIf type Is GetType(ArrayExpression) Then
                AnalyzeArrayExpression(CType(expression, ArrayExpression), leaveValueInStack, mustBeAssignable)

            ElseIf type Is GetType(InitializerExpression) Then
                AnalyzeInitializerExpression(CType(expression, InitializerExpression), leaveValueInStack, mustBeAssignable)

            ElseIf type Is GetType(IdentifierExpression) Then
                AnalyzeIdentifierExpression(CType(expression, IdentifierExpression), leaveValueInStack, mustBeAssignable)

            ElseIf type Is GetType(MethodCallExpression) Then
                AnalyzeMethodCallExpression(CType(expression, MethodCallExpression), leaveValueInStack, mustBeAssignable)

            ElseIf type Is GetType(NegativeExpression) Then
                AnalyzeNegativeExpression(CType(expression, NegativeExpression), leaveValueInStack, mustBeAssignable)

            ElseIf type Is GetType(PropertyExpression) Then
                AnalyzePropertyExpression(CType(expression, PropertyExpression), leaveValueInStack, mustBeAssignable)
            End If
        End Sub

        Private Sub AnalyzeStatement(statement As Statement)

            Dim type As Type = statement.GetType()

            If type Is GetType(AssignmentStatement) Then
                AnalyzeAssignmentStatement(CType(statement, AssignmentStatement))

            ElseIf type Is GetType(ElseIfStatement) Then
                AnalyzeElseIfStatement(CType(statement, ElseIfStatement))

            ElseIf type Is GetType(ForStatement) Then
                AnalyzeForStatement(CType(statement, ForStatement))

            ElseIf type Is GetType(ForeachStatement) Then
                AnalyzeForEachStatement(CType(statement, ForEachStatement))

            ElseIf type Is GetType(GotoStatement) Then
                AnalyzeGotoStatement(CType(statement, GotoStatement))

            ElseIf type Is GetType(IfStatement) Then
                AnalyzeIfStatement(CType(statement, IfStatement))

            ElseIf type Is GetType(MethodCallStatement) Then
                AnalyzeMethodCallStatement(CType(statement, MethodCallStatement))

            ElseIf type Is GetType(SubroutineCallStatement) Then
                AnalyzeSubroutineCallStatement(CType(statement, SubroutineCallStatement))

            ElseIf type Is GetType(SubroutineStatement) Then
                AnalyzeSubroutineStatement(CType(statement, SubroutineStatement))

            ElseIf type Is GetType(ReturnStatement) Then
                AnalyzeReturnStatement(CType(statement, ReturnStatement))

            ElseIf type Is GetType(WhileStatement) Then
                AnalyzeWhileStatement(CType(statement, WhileStatement))
            End If
        End Sub

        Private Sub AnalyzeBinaryExpression(binaryExpression As BinaryExpression, leaveValueInStack As Boolean, mustBeAssignable As Boolean)
            If binaryExpression.LeftHandSide IsNot Nothing Then
                AnalyzeExpression(binaryExpression.LeftHandSide, leaveValueInStack, mustBeAssignable)
            End If

            If binaryExpression.RightHandSide IsNot Nothing Then
                AnalyzeExpression(binaryExpression.RightHandSide, leaveValueInStack, mustBeAssignable)
            End If
        End Sub

        Private Sub AnalyzeArrayExpression(arrayExpression As ArrayExpression, leaveValueInStack As Boolean, mustBeAssignable As Boolean)
            If arrayExpression.LeftHand IsNot Nothing Then
                AnalyzeExpression(arrayExpression.LeftHand, leaveValueInStack, mustBeAssignable)
            End If

            If arrayExpression.Indexer IsNot Nothing Then
                AnalyzeExpression(arrayExpression.Indexer, leaveValueInStack, mustBeAssignable)
            End If
        End Sub

        Private Sub AnalyzeInitializerExpression(initExpression As InitializerExpression, leaveValueInStack As Boolean, mustBeAssignable As Boolean)
            If initExpression.Arguments IsNot Nothing Then
                For Each expr In initExpression.Arguments
                    AnalyzeExpression(expr, leaveValueInStack, mustBeAssignable)
                Next
            End If
        End Sub

        Private Sub AnalyzeReturnStatement(returnStatement As ReturnStatement)
            If returnStatement.ReturnExpression IsNot Nothing Then
                AnalyzeExpression(returnStatement.ReturnExpression, False, False)
            End If
        End Sub

        Private Sub AnalyzeIdentifierExpression(identifierExpression As IdentifierExpression, leaveValueInStack As Boolean, mustBeAssignable As Boolean)
            If identifierExpression.Identifier.Type <> 0 Then
                NoteVariableReference(identifierExpression.Identifier)

                If Not _symbolTable.IsDefined(identifierExpression) Then
                    Dim identifier = identifierExpression.Identifier
                    _symbolTable.Errors.Add(New [Error](identifier, $"The variable `{identifier.Text}` is used before being initialized."))
                End If
            End If
        End Sub

        Private Sub AnalyzeMethodCallExpression(methodCall As MethodCallExpression, leaveValueInStack As Boolean, mustBeAssignable As Boolean)
            NoteMethodCallReference(methodCall, leaveValueInStack, mustBeAssignable)

            If methodCall.TypeName.Type = TokenType.Illegal Then ' Function Call
                Dim subName = methodCall.MethodName.NormalizedText
                If Not _symbolTable.Subroutines.ContainsKey(subName) Then
                    _parser.AddError(methodCall.MethodName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("SubroutineNotDefined"), New Object(0) {methodCall.MethodName.Text}))
                Else
                    Dim pNo = GetParamNo(subName)
                    If pNo <> methodCall.Arguments.Count Then
                        _parser.AddError(methodCall.MethodName, $"`{methodCall.MethodName.Text}` expects {pNo} arguments.")
                    End If
                End If
            End If

            For Each argument In methodCall.Arguments
                AnalyzeExpression(argument, leaveValueInStack, mustBeAssignable)
            Next
        End Sub

        Private Sub AnalyzeNegativeExpression(negativeExpression As NegativeExpression, leaveValueInStack As Boolean, mustBeAssignable As Boolean)
            If negativeExpression.Expression IsNot Nothing Then
                AnalyzeExpression(negativeExpression.Expression, leaveValueInStack, mustBeAssignable)
            End If
        End Sub

        Private Sub AnalyzePropertyExpression(propertyExpression As PropertyExpression, leaveValueInStack As Boolean, mustBeAssignable As Boolean)
            NotePropertyReference(propertyExpression, leaveValueInStack, mustBeAssignable)
        End Sub

        Private Sub AnalyzeAssignmentStatement(assignmentStatement As AssignmentStatement)
            Dim idExpr = TryCast(assignmentStatement.RightValue, IdentifierExpression)

            If idExpr IsNot Nothing AndAlso _symbolTable.Subroutines.ContainsKey(idExpr.Identifier.NormalizedText) Then
                NoteEventReference(assignmentStatement.LeftValue, idExpr.Identifier)
                Return
            End If

            If assignmentStatement.LeftValue IsNot Nothing Then
                AnalyzeExpression(assignmentStatement.LeftValue, leaveValueInStack:=False, mustBeAssignable:=True)
            End If

            If assignmentStatement.RightValue IsNot Nothing Then
                AnalyzeExpression(assignmentStatement.RightValue, leaveValueInStack:=True, mustBeAssignable:=False)
            End If
        End Sub

        Private Sub AnalyzeElseIfStatement(elseIfStatement As ElseIfStatement)
            If elseIfStatement.Condition IsNot Nothing Then
                AnalyzeExpression(elseIfStatement.Condition, leaveValueInStack:=True, mustBeAssignable:=False)
            End If

            For Each thenStatement In elseIfStatement.ThenStatements
                AnalyzeStatement(thenStatement)
            Next
        End Sub

        Private Sub AnalyzeForStatement(forStatement As ForStatement)
            If forStatement.Iterator.Type <> 0 Then
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

            For Each item In forStatement.Body
                AnalyzeStatement(item)
            Next
        End Sub

        Private Sub AnalyzeForEachStatement(forEach As ForEachStatement)
            If forEach.Iterator.Type <> 0 Then
                NoteVariableReference(forEach.Iterator)
            End If


            If forEach.ArrayExpression IsNot Nothing Then
                AnalyzeExpression(forEach.ArrayExpression, leaveValueInStack:=True, mustBeAssignable:=False)
            End If

            For Each item In forEach.Body
                AnalyzeStatement(item)
            Next
        End Sub


        Private Sub AnalyzeGotoStatement(gotoStatement As GotoStatement)
            Dim label = gotoStatement.Label
            If label.Type <> 0 Then
                If label.Type <> 0 AndAlso Not _symbolTable.Labels.ContainsKey(label.NormalizedText) Then
                    _parser.AddError(label, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("LabelNotFound"), New Object(0) {label.Text}))
                Else
                    Dim labelStatement = CType(_symbolTable.Labels(label.NormalizedText).Parent, LabelStatement)
                    If labelStatement.subroutine?.Name.NormalizedText <> gotoStatement.subroutine?.Name.NormalizedText Then
                        _parser.SymbolTable.Errors.Add(New [Error](label, "GoTo can't jump accross subroutines."))
                    End If
                End If
            End If
        End Sub

        Private Sub AnalyzeIfStatement(ifStatement As IfStatement)
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

        Private Sub AnalyzeMethodCallStatement(methodCallStatement As MethodCallStatement)
            If methodCallStatement.MethodCallExpression IsNot Nothing Then
                AnalyzeExpression(methodCallStatement.MethodCallExpression, leaveValueInStack:=False, mustBeAssignable:=False)
            End If
        End Sub

        Function GetParamNo(subName As String)
            For Each statement In _parser.ParseTree
                Dim subroutine = TryCast(statement, SubroutineStatement)
                If subroutine IsNot Nothing AndAlso subroutine.Name.NormalizedText = subName Then
                    Return subroutine.Params?.Count
                End If
            Next
            Return 0
        End Function

        Private Sub AnalyzeSubroutineCallStatement(subroutineCall As SubroutineCallStatement)
            Dim subroutineName = subroutineCall.Name
            If subroutineName.Type <> 0 Then
                Dim subName = subroutineName.NormalizedText
                If Not _symbolTable.Subroutines.ContainsKey(subName) Then
                    _parser.AddError(subroutineName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("SubroutineNotDefined"), New Object(0) {subroutineName.Text}))
                Else
                    Dim pNo = GetParamNo(subName)
                    If pNo <> subroutineCall.Args.Count Then
                        _parser.AddError(subroutineName, $"`{subroutineName.Text}` expects {pNo} arguments.")
                    End If
                End If

                For Each arg In subroutineCall.Args
                    AnalyzeExpression(arg, False, False)
                Next
            End If
        End Sub

        Private Sub AnalyzeSubroutineStatement(subroutine As SubroutineStatement)

            If subroutine.Name.Text = "_" Then
                _parser.AddError(subroutine.Name, "_ is not a valid name")
            End If

            Select Case subroutine.StartToken.Type
                Case TokenType.Sub
                    If subroutine.EndSubToken.Type <> TokenType.EndSub Then
                        _parser.AddError(subroutine.EndSubToken, "Sub must end with EndSub")
                    End If

                Case Else
                    If subroutine.EndSubToken.Type <> TokenType.EndFunction Then
                        _parser.AddError(subroutine.EndSubToken, "Function must end with EndFunction")
                    End If

                    If subroutine.ReturnStatements.Count = 0 Then
                        _parser.AddError(subroutine.SubToken, "Function must return a value")
                    End If
            End Select

            For Each item In subroutine.Body
                AnalyzeStatement(item)
            Next
        End Sub

        Private Sub AnalyzeWhileStatement(whileStatement As WhileStatement)
            If whileStatement.Condition IsNot Nothing Then
                AnalyzeExpression(whileStatement.Condition, leaveValueInStack:=True, mustBeAssignable:=False)
            End If

            For Each item In whileStatement.Body
                AnalyzeStatement(item)
            Next
        End Sub

        Private Sub NoteEventReference(leftValue As Expression, subroutineName As Token)
            Dim propertyExpression = TryCast(leftValue, PropertyExpression)

            If propertyExpression IsNot Nothing Then
                Dim typeName = propertyExpression.TypeName
                Dim propertyName = propertyExpression.PropertyName

                If typeName.Type = TokenType.Illegal OrElse propertyName.Type = TokenType.Illegal Then
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

        Private Sub NoteMethodCallReference(methodExpression As MethodCallExpression, leaveValueInStack As Boolean, isAssignable As Boolean)
            Dim typeNameInfo = methodExpression.TypeName
            Dim methodName = methodExpression.MethodName

            If typeNameInfo.Type = TokenType.Illegal OrElse methodName.Type = TokenType.Illegal Then
                Return
            End If

            Dim value As TypeInfo = Nothing
            Dim typeName = typeNameInfo.NormalizedText

            If _typeInfoBag.Types.TryGetValue(typeName, value) Then
                Dim value2 As MethodInfo = Nothing

                If value.Methods.TryGetValue(methodName.NormalizedText, value2) Then
                    Dim num As Integer = value2.GetParameters().Length

                    If num <> methodExpression.Arguments.Count Then
                        _parser.AddError(methodName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("ArgumentNumberMismatch"), typeNameInfo.Text, methodName.Text, methodExpression.Arguments.Count, num))
                    End If

                    If leaveValueInStack AndAlso value2.ReturnType Is GetType(Void) Then
                        Me._parser.AddError(methodName, String.Format(System.Globalization.CultureInfo.CurrentUICulture, Microsoft.SmallBasic.ResourceHelper.GetString("ReturnValueExpectedFromVoidMethod"),
                                New Object() {typeNameInfo.Text, methodName.Text}))
                    End If
                Else
                    _parser.AddError(methodName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("MethodNotFound"), New Object(1) {methodName.Text, typeNameInfo.Text}))
                End If
            Else
                _parser.AddError(typeNameInfo, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("TypeNotFound"), New Object(0) {typeNameInfo.Text}))
            End If
        End Sub

        Private Sub NotePropertyReference(propExpr As PropertyExpression, leaveValueInStack As Boolean, mustBeAssignable As Boolean)
            Dim typeNameInfo = propExpr.TypeName
            Dim propertyNameInfo = propExpr.PropertyName

            If typeNameInfo.Type = TokenType.Illegal OrElse propertyNameInfo.Type = TokenType.Illegal Then
                Return
            End If

            Dim value As TypeInfo = Nothing
            Dim typeName = typeNameInfo.NormalizedText
            Dim propertyName = propertyNameInfo.NormalizedText

            If propExpr.IsDynamic Then
                Dim subroutine = SubroutineStatement.Current
                If subroutine Is Nothing Then
                    subroutine = SubroutineStatement.GetSubroutine(propExpr)
                End If

                Dim id = New IdentifierExpression() With {
                        .Identifier = typeNameInfo,
                        .Subroutine = subroutine
                }

                If Not _symbolTable.IsDefined(id) Then
                    _symbolTable.Errors.Add(New [Error](typeNameInfo, $"The variable `{typeNameInfo.Text}` is used before being initialized."))

                ElseIf _symbolTable.Dynamics.ContainsKey(typeName) Then
                    If Not _symbolTable.Dynamics(typeName).ContainsKey(propertyName) Then
                        _symbolTable.Errors.Add(New [Error](propertyNameInfo, "Trying to read property before setting its value. "))
                    End If
                End If

            ElseIf _typeInfoBag.Types.TryGetValue(typeName, value) Then
                Dim prop As PropertyInfo = Nothing
                Dim ev As EventInfo = Nothing

                If value.Properties.TryGetValue(propertyName, prop) Then
                    If mustBeAssignable AndAlso Not prop.CanWrite Then
                        _parser.AddError(propertyNameInfo, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("PropertyIsReadOnly"), New Object(1) {propertyNameInfo.Text, typeNameInfo.Text}))
                    End If

                    If leaveValueInStack AndAlso Not prop.CanRead Then
                        _parser.AddError(propertyNameInfo, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("PropertyIsWriteOnly"), New Object(1) {propertyNameInfo.Text, typeNameInfo.Text}))
                    End If

                ElseIf value.Events.TryGetValue(propertyNameInfo.NormalizedText, ev) Then
                    _parser.AddError(propertyNameInfo, $"Event {ev.Name} can only be set to a Sub.")
                Else
                    _parser.AddError(propertyNameInfo, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("PropertyNotFound"), New Object(1) {propertyNameInfo.Text, typeNameInfo.Text}))
                End If
            Else
                _parser.AddError(typeNameInfo, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("TypeNotFound"), New Object(0) {typeNameInfo.Text}))
            End If
        End Sub


        Private Sub NoteVariableReference(variable As Token)
            If variable.Type <> 0 AndAlso Not _symbolTable.GlobalVariables.ContainsKey(variable.NormalizedText) AndAlso
                       _symbolTable.Subroutines.ContainsKey(variable.NormalizedText) Then
                _parser.AddError(variable, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("SubroutineUsedAsVariable"), New Object(0) {variable.Text}))
            End If
        End Sub
    End Class
End Namespace
