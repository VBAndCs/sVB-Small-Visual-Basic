Imports System
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Threading
Imports Microsoft.SmallVisualBasic.Expressions
Imports Microsoft.SmallVisualBasic.Library

Namespace Microsoft.SmallVisualBasic.Engine
    Public Class ProgramEngine
        Inherits MarshalByRefObject

        Private _DebuggerAppDomain As AppDomain
        Private _Compiler As Compiler
        Private _CurrentDebuggerState As DebuggerState
        Private _CurrentInstruction As Instruction
        Private _Translator As ProgramTranslator

        <Serializable>
        Private Class ProgramRunner
            Private previousLineNumber As Integer = -1

            Public Property Breakpoints As List(Of Integer)

            Public Property CurrentInstruction As Instruction

            Public Property DebuggerCommand As DebuggerCommand

            Public Property DebuggerExecution As ManualResetEvent

            Public Property DebuggerState As DebuggerState

            Public Property Instructions As List(Of Instruction)

            Public Property Fields As Dictionary(Of String, Primitive)

            Public Property LabelMap As Dictionary(Of String, Integer)

            Public Property SubroutineInstructions As Dictionary(Of String, List(Of Instruction))

            Public Property TypeInfoBag As TypeInfoBag

            Public Event LineNumberChanged As EventHandler
            Public Event DebuggerStateChanged As EventHandler

            Public Sub New()
                Breakpoints = New List(Of Integer)()
                TypeInfoBag = TryCast(AppDomain.CurrentDomain.GetData("TypeInfoBag"), TypeInfoBag)
                Instructions = TryCast(AppDomain.CurrentDomain.GetData("ProgramInstructions"), List(Of Instruction))
                SubroutineInstructions = TryCast(AppDomain.CurrentDomain.GetData("SubroutineInstructions"), Dictionary(Of String, List(Of Instruction)))
                Fields = New Dictionary(Of String, Primitive)()
                LabelMap = New Dictionary(Of String, Integer)()
                DebuggerExecution = New ManualResetEvent(initialState:=True)
            End Sub

            Private Sub ChangeDebuggerState(state As DebuggerState)
                If DebuggerState <> state Then
                    DebuggerState = state
                    RaiseEvent DebuggerStateChanged(Me, EventArgs.Empty)
                End If
            End Sub

            Private Sub CheckForExecutionBreak()
                If CurrentInstruction.LineNumber <> previousLineNumber Then
                    If Breakpoints.Contains(CurrentInstruction.LineNumber) Then
                        DebuggerExecution.Reset()
                    End If

                    If Not DebuggerExecution.WaitOne(0) Then
                        ChangeDebuggerState(DebuggerState.Paused)
                        DebuggerExecution.WaitOne()
                    End If
                End If

                ChangeDebuggerState(DebuggerState.Running)
            End Sub

            Private Sub PrepareDebuggerForNextInstruction()
                If CurrentInstruction.LineNumber <> previousLineNumber AndAlso (
                            DebuggerCommand = DebuggerCommand.StepInto OrElse
                            DebuggerCommand = DebuggerCommand.StepOver
                        ) Then
                    DebuggerExecution.Reset()
                End If
            End Sub

            Public Sub [Continue]()
                DebuggerExecution.Set()
            End Sub

            Public Sub Pause()
                DebuggerExecution.Reset()
            End Sub

            Public Sub Reset()
                Fields.Clear()
            End Sub

            Public Sub StepInto()
                DebuggerCommand = DebuggerCommand.StepInto
                [Continue]()
            End Sub

            Public Sub StepOver()
                DebuggerCommand = DebuggerCommand.StepOver
                [Continue]()
            End Sub

            Public Sub RunProgram(stopOnFirstInstruction As Boolean)
                If stopOnFirstInstruction Then
                    DebuggerExecution.Reset()
                End If

                Dim thread As Thread = New Thread(Sub() ExecuteInstructions(Instructions))
                thread.IsBackground = True
                thread.Start()
            End Sub

            Private Sub ExecuteInstructions(instructions As List(Of Instruction))
                For i = 0 To instructions.Count - 1
                    Dim labelInstruction = TryCast(instructions(i), LabelInstruction)

                    If labelInstruction IsNot Nothing Then
                        LabelMap(labelInstruction.LabelName) = i
                    End If
                Next

                Dim num = 0

                While num < instructions.Count
                    CurrentInstruction = instructions(num)

                    If CurrentInstruction.LineNumber <> previousLineNumber AndAlso LineNumberChangedEvent IsNot Nothing Then
                        RaiseEvent LineNumberChanged(Me, EventArgs.Empty)
                    End If

                    CheckForExecutionBreak()
                    Dim text = ExecuteInstruction(CurrentInstruction)
                    PrepareDebuggerForNextInstruction()
                    previousLineNumber = CurrentInstruction.LineNumber
                    num = If(Not Equals(text, Nothing), LabelMap(text), num + 1)
                End While

                ChangeDebuggerState(DebuggerState.Finished)
            End Sub

            Private Function ExecuteInstruction(instruction As Instruction) As String
                Select Case instruction.InstructionType
                    Case InstructionType.ArrayAssignment
                        Return ExecuteArrayAssignmentInstruction(TryCast(instruction, ArrayAssignmentInstruction))
                    Case InstructionType.EventAssignment
                        Return ExecuteEventAssignmentInstruction(TryCast(instruction, EventAssignmentInstruction))
                    Case InstructionType.FieldAssignment
                        Return ExecuteFieldAssignmentInstruction(TryCast(instruction, FieldAssignmentInstruction))
                    Case InstructionType.Goto
                        Return CType(instruction, GotoInstruction).LabelName
                    Case InstructionType.IfGoto
                        Return ExecuteIfGotoInstruction(TryCast(instruction, IfGotoInstruction))
                    Case InstructionType.IfNotGoto
                        Return ExecuteIfNotGotoInstruction(TryCast(instruction, IfNotGotoInstruction))
                    Case InstructionType.LabelInstruction
                        Return Nothing
                    Case InstructionType.MethodCall
                        Return EvaluateMethodCallInstruction(TryCast(instruction, MethodCallInstruction))
                    Case InstructionType.PropertyAssignment
                        Return ExecutePropertyAssignmentInstruction(TryCast(instruction, PropertyAssignmentInstruction))
                    Case InstructionType.SubroutineCall
                        Return ExecuteSubroutineCallInstruction(TryCast(instruction, SubroutineCallInstruction))
                    Case Else
                        Return Nothing
                End Select
            End Function

            Private Function ExecuteArrayAssignmentInstruction(instruction As ArrayAssignmentInstruction) As String
                SetArrayValue(instruction.ArrayExpression, EvaluateExpression(instruction.RightSide))
                Return Nothing
            End Function

            Private Sub SetArrayValue(lvalue As ArrayExpression, value As Primitive)
                Dim identifierExpression = TryCast(lvalue.LeftHand, IdentifierExpression)
                Dim value2 As Primitive = Nothing

                If identifierExpression IsNot Nothing Then
                    Dim normalizedText = identifierExpression.Identifier.LCaseText

                    If Not Fields.TryGetValue(normalizedText, value2) Then
                        value2 = Nothing
                    End If

                    Fields(normalizedText) = Primitive.SetArrayValue(value, value2, EvaluateExpression(lvalue.Indexer))
                Else
                    Dim arrayExpression = TryCast(lvalue.LeftHand, ArrayExpression)

                    If arrayExpression IsNot Nothing Then
                        SetArrayValue(arrayExpression, Primitive.SetArrayValue(value, EvaluateArrayExpression(arrayExpression), EvaluateExpression(lvalue.Indexer)))
                    End If
                End If
            End Sub

            Private Function ExecuteEventAssignmentInstruction(instruction As EventAssignmentInstruction) As String
                instruction.EventInfo.AddEventHandler(Nothing, CType(Sub() ExecuteInstructions(SubroutineInstructions(instruction.SubroutineName)), SmallVisualBasicCallback))
                Return Nothing
            End Function

            Private Function ExecuteFieldAssignmentInstruction(instruction As FieldAssignmentInstruction) As String
                Fields(instruction.FieldName) = EvaluateExpression(instruction.RightSide)
                Return Nothing
            End Function

            Private Function ExecuteIfGotoInstruction(instruction As IfGotoInstruction) As String
                If Primitive.ConvertToBoolean(EvaluateExpression(instruction.Condition)) Then
                    Return instruction.LabelName
                End If

                Return Nothing
            End Function

            Private Function ExecuteIfNotGotoInstruction(instruction As IfNotGotoInstruction) As String
                If Not Primitive.ConvertToBoolean(EvaluateExpression(instruction.Condition)) Then
                    Return instruction.LabelName
                End If

                Return Nothing
            End Function

            Private Function EvaluateMethodCallInstruction(instruction As MethodCallInstruction) As String
                EvaluateMethodCallExpression(instruction.MethodExpression)
                Return Nothing
            End Function

            Private Function ExecutePropertyAssignmentInstruction(instruction As PropertyAssignmentInstruction) As String
                instruction.PropertyInfo.SetValue(Nothing, EvaluateExpression(instruction.RightSide), Nothing)
                Return Nothing
            End Function

            Private Function ExecuteSubroutineCallInstruction(instruction As SubroutineCallInstruction) As String
                Dim debuggerCommand = Me.DebuggerCommand

                If Me.DebuggerCommand = DebuggerCommand.StepOver Then
                    Me.DebuggerCommand = DebuggerCommand.Run
                ElseIf Me.DebuggerCommand = DebuggerCommand.StepInto Then
                    Pause()
                End If

                ExecuteInstructions(SubroutineInstructions(instruction.SubroutineName))
                Me.DebuggerCommand = debuggerCommand
                Return Nothing
            End Function

            Private Function EvaluateExpression(expression As Expression) As Primitive
                If TypeOf expression Is ArrayExpression Then
                    Return EvaluateArrayExpression(TryCast(expression, ArrayExpression))
                End If

                If TypeOf expression Is BinaryExpression Then
                    Return EvaluateBinaryExpression(TryCast(expression, BinaryExpression))
                End If

                If TypeOf expression Is IdentifierExpression Then
                    Return EvaluateIdentifierExpression(TryCast(expression, IdentifierExpression))
                End If

                If TypeOf expression Is LiteralExpression Then
                    Return EvaluateLiteralExpression(TryCast(expression, LiteralExpression))
                End If

                If TypeOf expression Is MethodCallExpression Then
                    Return EvaluateMethodCallExpression(TryCast(expression, MethodCallExpression))
                End If

                If TypeOf expression Is PropertyExpression Then
                    Return EvaluatePropertyExpression(TryCast(expression, PropertyExpression))
                End If

                If TypeOf expression Is NegativeExpression Then
                    Return EvaluateNegativeExpression(TryCast(expression, NegativeExpression))
                End If

                Return Nothing
            End Function

            Private Function EvaluateArrayExpression(expression As ArrayExpression) As Primitive
                Dim identifierExpression As IdentifierExpression = TryCast(expression.LeftHand, IdentifierExpression)
                Dim value As Primitive = Nothing

                If identifierExpression IsNot Nothing Then
                    If Not Fields.TryGetValue(identifierExpression.Identifier.LCaseText, value) Then
                        value = ""
                    End If

                    Return Primitive.GetArrayValue(value, EvaluateExpression(expression.Indexer))
                End If

                Dim arrayExpression As ArrayExpression = TryCast(expression.LeftHand, ArrayExpression)

                If arrayExpression IsNot Nothing Then
                    Return Primitive.GetArrayValue(EvaluateArrayExpression(arrayExpression), EvaluateExpression(expression.Indexer))
                End If

                Return Nothing
            End Function

            Private Function EvaluateBinaryExpression(expression As BinaryExpression) As Primitive
                Dim leftExpr = EvaluateExpression(expression.LeftHandSide)
                Dim rightExpr = EvaluateExpression(expression.RightHandSide)

                Select Case expression.Operator.Type
                    Case TokenType.Concatenation
                        Return leftExpr.Concat(rightExpr)
                    Case TokenType.Addition
                        Return leftExpr.Add(rightExpr)
                    Case TokenType.And
                        Return Primitive.op_And(leftExpr, rightExpr)
                    Case TokenType.Division
                        Return leftExpr.Divide(rightExpr)
                    Case TokenType.Equals
                        Return leftExpr.EqualTo(rightExpr)
                    Case TokenType.GreaterThan
                        Return leftExpr.GreaterThan(rightExpr)
                    Case TokenType.GreaterThanEqualTo
                        Return leftExpr.GreaterThanOrEqualTo(rightExpr)
                    Case TokenType.LessThan
                        Return leftExpr.LessThan(rightExpr)
                    Case TokenType.LessThanEqualTo
                        Return leftExpr.LessThanOrEqualTo(rightExpr)
                    Case TokenType.Multiplication
                        Return leftExpr.Multiply(rightExpr)
                    Case TokenType.NotEqualTo
                        Return leftExpr.NotEqualTo(rightExpr)
                    Case TokenType.Or
                        Return Primitive.op_Or(leftExpr, rightExpr)
                    Case TokenType.Subtraction
                        Return leftExpr.Subtract(rightExpr)
                    Case Else
                        Return Nothing
                End Select
            End Function

            Private Function EvaluateIdentifierExpression(expression As IdentifierExpression) As Primitive
                Dim value As Primitive = Nothing

                If Fields.TryGetValue(expression.Identifier.LCaseText, value) Then
                    Return value
                End If

                Return Nothing
            End Function

            Private Function EvaluateLiteralExpression(expression As LiteralExpression) As Primitive
                If expression.Literal.Type = TokenType.StringLiteral Then
                    Return New Primitive(expression.Literal.Text.Trim(""""c))
                End If

                Return New Primitive(expression.Literal.Text)
            End Function

            Private Function EvaluateMethodCallExpression(expression As MethodCallExpression) As Primitive
                Dim typeInfo = TypeInfoBag.Types(expression.TypeName.LCaseText)
                Dim methodInfo = typeInfo.Methods(expression.MethodName.LCaseText)
                Dim list As List(Of Object) = New List(Of Object)()

                For Each argument In expression.Arguments
                    list.Add(EvaluateExpression(argument))
                Next

                Dim obj As Object = methodInfo.Invoke(Nothing, list.ToArray())
                Return If(TypeOf obj Is Primitive, obj, Nothing)
            End Function

            Private Function EvaluateNegativeExpression(expression As NegativeExpression) As Primitive
                Return -EvaluateExpression(expression.Expression)
            End Function

            Private Function EvaluatePropertyExpression(expression As PropertyExpression) As Primitive
                Dim typeInfo = TypeInfoBag.Types(expression.TypeName.LCaseText)
                Dim propertyInfo = typeInfo.Properties(expression.PropertyName.LCaseText)
                Dim value = propertyInfo.GetValue(Nothing, Nothing)

                If TypeOf value Is Primitive Then
                    Return value
                End If

                Return Nothing
            End Function
        End Class

        Public ReadOnly Property Breakpoints As List(Of Integer)
            Get
                Return Engine.Breakpoints
            End Get
        End Property

        Public Property DebuggerAppDomain As AppDomain
            Get
                Return _DebuggerAppDomain
            End Get
            Private Set(value As AppDomain)
                _DebuggerAppDomain = value
            End Set
        End Property

        Public Property Compiler As Compiler
            Get
                Return _Compiler
            End Get
            Private Set(value As Compiler)
                _Compiler = value
            End Set
        End Property

        Public Property CurrentDebuggerState As DebuggerState
            Get
                Return _CurrentDebuggerState
            End Get
            Private Set(value As DebuggerState)
                _CurrentDebuggerState = value
            End Set
        End Property

        Public Property CurrentInstruction As Instruction
            Get
                Return _CurrentInstruction
            End Get
            Private Set(value As Instruction)
                _CurrentInstruction = value
            End Set
        End Property

        Public Shared ReadOnly Property Engine As ProgramEngine
            Get
                Return TryCast(AppDomain.CurrentDomain.GetData("ProgramEngine"), ProgramEngine)
            End Get
        End Property

        Public Property Translator As ProgramTranslator
            Get
                Return _Translator
            End Get
            Private Set(value As ProgramTranslator)
                _Translator = value
            End Set
        End Property

        Private Shared Property Runner As ProgramRunner
        Public Event DebuggerStateChanged As EventHandler
        Public Event EngineUnloaded As EventHandler
        Public Event LineNumberChanged As EventHandler

        Public Sub New(compiler As Compiler)
            Me.Compiler = compiler
            Translator = New ProgramTranslator(Me.Compiler)
            Translator.TranslateProgram()
            Dim info As AppDomainSetup = New AppDomainSetup With {
                .LoaderOptimization = LoaderOptimization.MultiDomain
            }
            Me.DebuggerAppDomain = AppDomain.CreateDomain("Debuggee", Nothing, info)
            Me.DebuggerAppDomain.SetData("TypeInfoBag", Me.Compiler.TypeInfoBag)
            Me.DebuggerAppDomain.SetData("ProgramInstructions", Translator.ProgramInstructions)
            Me.DebuggerAppDomain.SetData("SubroutineInstructions", Translator.SubroutineInstructions)
            Me.DebuggerAppDomain.SetData("ProgramEngine", Me)
            Me.DebuggerAppDomain.DoCallBack(New CrossAppDomainDelegate(AddressOf InitializeRunner))
            Dim debuggerAppDomain = Me.DebuggerAppDomain
            Dim value As EventHandler = Sub() RaiseEvent EngineUnloaded(Me, EventArgs.Empty)
            AddHandler debuggerAppDomain.DomainUnload, value
        End Sub

        Private Shared Sub InitializeRunner()
            Runner = New ProgramRunner()
            AddHandler Runner.DebuggerStateChanged, Sub() Engine.RaiseDebuggerStateChangedEvent(Runner.DebuggerState)
            AddHandler Runner.LineNumberChanged, Sub() Engine.RaiseLineNumberChangedEvent(Runner.CurrentInstruction)
        End Sub

        Public Sub AddBreakpoint(lineNumber As Integer)
            DebuggerAppDomain.DoCallBack(Sub() Runner.Breakpoints.Add(lineNumber))
        End Sub

        Public Sub [Continue]()
            DebuggerAppDomain.DoCallBack(Sub() Runner.Continue())
        End Sub

        Public Sub Pause()
            DebuggerAppDomain.DoCallBack(Sub() Runner.Pause())
        End Sub

        Private Sub RaiseDebuggerStateChangedEvent(currentState As DebuggerState)
            CurrentDebuggerState = currentState
            RaiseEvent DebuggerStateChanged(Me, EventArgs.Empty)
        End Sub

        Private Sub RaiseLineNumberChangedEvent(currentInstruction As Instruction)
            Me.CurrentInstruction = currentInstruction
            RaiseEvent LineNumberChanged(Me, EventArgs.Empty)
        End Sub

        Public Sub StepInto()
            DebuggerAppDomain.DoCallBack(Sub() Runner.StepInto())
        End Sub

        Public Sub StepOver()
            DebuggerAppDomain.DoCallBack(Sub() Runner.StepOver())
        End Sub

        Public Sub Reset()
            DebuggerAppDomain.DoCallBack(Sub() Runner.Reset())
        End Sub

        Public Sub RunProgram()
            DebuggerAppDomain.DoCallBack(Sub() Runner.RunProgram(stopOnFirstInstruction:=False))
        End Sub

        Public Sub RunProgram(stopOnFirstInstruction As Boolean)
            DebuggerAppDomain.DoCallBack(Sub() Runner.RunProgram(stopOnFirstInstruction:=True))
        End Sub
    End Class
End Namespace
