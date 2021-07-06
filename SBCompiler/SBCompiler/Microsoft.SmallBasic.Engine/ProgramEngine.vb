Imports System
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Threading
Imports Microsoft.SmallBasic.Expressions
Imports SmallBasicLibrary.Microsoft.SmallBasic.Library

Namespace Microsoft.SmallBasic.Engine
    Public Class ProgramEngine
        Inherits MarshalByRefObject

        Private _DebuggerAppDomain As AppDomain
        Private _Compiler As Compiler
        Private _CurrentDebuggerState As DebuggerState
        Private _CurrentInstruction As Instruction
        Private _Translator As ProgramTranslator

        <Serializable>
        Private Class ProgramRunner
            Private _Breakpoints As List(Of Integer)
            Private _CurrentInstruction As Instruction
            Private _DebuggerCommand As DebuggerCommand
            Private _DebuggerExecution As ManualResetEvent
            Private _DebuggerState As DebuggerState
            Private _Instructions As List(Of Instruction)
            Private _Fields As Dictionary(Of String, Primitive)
            Private _LabelMap As Dictionary(Of String, Integer)
            Private _SubroutineInstructions As Dictionary(Of String, List(Of Instruction))
            Private _TypeInfoBag As TypeInfoBag
            Private previousLineNumber As Integer = -1

            Public Property Breakpoints As List(Of Integer)
                Get
                    Return _Breakpoints
                End Get
                Private Set(ByVal value As List(Of Integer))
                    _Breakpoints = value
                End Set
            End Property

            Public Property CurrentInstruction As Instruction
                Get
                    Return _CurrentInstruction
                End Get
                Private Set(ByVal value As Instruction)
                    _CurrentInstruction = value
                End Set
            End Property

            Public Property DebuggerCommand As DebuggerCommand
                Get
                    Return _DebuggerCommand
                End Get
                Private Set(ByVal value As DebuggerCommand)
                    _DebuggerCommand = value
                End Set
            End Property

            Public Property DebuggerExecution As ManualResetEvent
                Get
                    Return _DebuggerExecution
                End Get
                Private Set(ByVal value As ManualResetEvent)
                    _DebuggerExecution = value
                End Set
            End Property

            Public Property DebuggerState As DebuggerState
                Get
                    Return _DebuggerState
                End Get
                Private Set(ByVal value As DebuggerState)
                    _DebuggerState = value
                End Set
            End Property

            Public Property Instructions As List(Of Instruction)
                Get
                    Return _Instructions
                End Get
                Private Set(ByVal value As List(Of Instruction))
                    _Instructions = value
                End Set
            End Property

            Public Property Fields As Dictionary(Of String, Primitive)
                Get
                    Return _Fields
                End Get
                Private Set(ByVal value As Dictionary(Of String, Primitive))
                    _Fields = value
                End Set
            End Property

            Public Property LabelMap As Dictionary(Of String, Integer)
                Get
                    Return _LabelMap
                End Get
                Private Set(ByVal value As Dictionary(Of String, Integer))
                    _LabelMap = value
                End Set
            End Property

            Public Property SubroutineInstructions As Dictionary(Of String, List(Of Instruction))
                Get
                    Return _SubroutineInstructions
                End Get
                Private Set(ByVal value As Dictionary(Of String, List(Of Instruction)))
                    _SubroutineInstructions = value
                End Set
            End Property

            Public Property TypeInfoBag As TypeInfoBag
                Get
                    Return _TypeInfoBag
                End Get
                Private Set(ByVal value As TypeInfoBag)
                    _TypeInfoBag = value
                End Set
            End Property

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

            Private Sub ChangeDebuggerState(ByVal state As DebuggerState)
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
                If CurrentInstruction.LineNumber <> previousLineNumber AndAlso (DebuggerCommand = DebuggerCommand.StepInto OrElse DebuggerCommand = DebuggerCommand.StepOver) Then
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

            Public Sub RunProgram(ByVal stopOnFirstInstruction As Boolean)
                If stopOnFirstInstruction Then
                    DebuggerExecution.Reset()
                End If

                Dim thread As Thread = New Thread(Sub() ExecuteInstructions(Instructions))
                thread.IsBackground = True
                thread.Start()
            End Sub

            Private Sub ExecuteInstructions(ByVal instructions As List(Of Instruction))
                For i = 0 To instructions.Count - 1
                    Dim labelInstruction As LabelInstruction = TryCast(instructions(i), LabelInstruction)

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

            Private Function ExecuteInstruction(ByVal instruction As Instruction) As String
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

            Private Function ExecuteArrayAssignmentInstruction(ByVal instruction As ArrayAssignmentInstruction) As String
                SetArrayValue(instruction.ArrayExpression, EvaluateExpression(instruction.RightSide))
                Return Nothing
            End Function

            Private Sub SetArrayValue(ByVal lvalue As ArrayExpression, ByVal value As Primitive)
                Dim identifierExpression As IdentifierExpression = TryCast(lvalue.LeftHand, IdentifierExpression)
                Dim value2 As Primitive = Nothing

                If identifierExpression IsNot Nothing Then
                    Dim normalizedText = identifierExpression.Identifier.NormalizedText

                    If Not Fields.TryGetValue(normalizedText, value2) Then
                        value2 = Nothing
                    End If

                    Fields(normalizedText) = Primitive.SetArrayValue(value, value2, EvaluateExpression(lvalue.Indexer))
                Else
                    Dim arrayExpression As ArrayExpression = TryCast(lvalue.LeftHand, ArrayExpression)

                    If arrayExpression IsNot Nothing Then
                        SetArrayValue(arrayExpression, Primitive.SetArrayValue(value, EvaluateArrayExpression(arrayExpression), EvaluateExpression(lvalue.Indexer)))
                    End If
                End If
            End Sub

            Private Function ExecuteEventAssignmentInstruction(ByVal instruction As EventAssignmentInstruction) As String
                instruction.EventInfo.AddEventHandler(Nothing, CType(Sub() ExecuteInstructions(SubroutineInstructions(instruction.SubroutineName)), SmallBasicCallback))
                Return Nothing
            End Function

            Private Function ExecuteFieldAssignmentInstruction(ByVal instruction As FieldAssignmentInstruction) As String
                Fields(instruction.FieldName) = EvaluateExpression(instruction.RightSide)
                Return Nothing
            End Function

            Private Function ExecuteIfGotoInstruction(ByVal instruction As IfGotoInstruction) As String
                If Primitive.ConvertToBoolean(EvaluateExpression(instruction.Condition)) Then
                    Return instruction.LabelName
                End If

                Return Nothing
            End Function

            Private Function ExecuteIfNotGotoInstruction(ByVal instruction As IfNotGotoInstruction) As String
                If Not Primitive.ConvertToBoolean(EvaluateExpression(instruction.Condition)) Then
                    Return instruction.LabelName
                End If

                Return Nothing
            End Function

            Private Function EvaluateMethodCallInstruction(ByVal instruction As MethodCallInstruction) As String
                EvaluateMethodCallExpression(instruction.MethodExpression)
                Return Nothing
            End Function

            Private Function ExecutePropertyAssignmentInstruction(ByVal instruction As PropertyAssignmentInstruction) As String
                instruction.PropertyInfo.SetValue(Nothing, EvaluateExpression(instruction.RightSide), Nothing)
                Return Nothing
            End Function

            Private Function ExecuteSubroutineCallInstruction(ByVal instruction As SubroutineCallInstruction) As String
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

            Private Function EvaluateExpression(ByVal expression As Expression) As Primitive
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

            Private Function EvaluateArrayExpression(ByVal expression As ArrayExpression) As Primitive
                Dim identifierExpression As IdentifierExpression = TryCast(expression.LeftHand, IdentifierExpression)
                Dim value As Primitive = Nothing

                If identifierExpression IsNot Nothing Then
                    If Not Fields.TryGetValue(identifierExpression.Identifier.NormalizedText, value) Then
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

            Private Function EvaluateBinaryExpression(ByVal expression As BinaryExpression) As Primitive
                Dim primitive = EvaluateExpression(expression.LeftHandSide)
                Dim primitive2 = EvaluateExpression(expression.RightHandSide)

                Select Case expression.Operator.Token
                    Case Token.Addition
                        Return primitive.Add(primitive2)
                    Case Token.And
                        Return Primitive.op_And(primitive, primitive2)
                    Case Token.Division
                        Return primitive.Divide(primitive2)
                    Case Token.Equals
                        Return primitive.EqualTo(primitive2)
                    Case Token.GreaterThan
                        Return primitive.GreaterThan(primitive2)
                    Case Token.GreaterThanEqualTo
                        Return primitive.GreaterThanOrEqualTo(primitive2)
                    Case Token.LessThan
                        Return primitive.LessThan(primitive2)
                    Case Token.LessThanEqualTo
                        Return primitive.LessThanOrEqualTo(primitive2)
                    Case Token.Multiplication
                        Return primitive.Multiply(primitive2)
                    Case Token.NotEqualTo
                        Return primitive.NotEqualTo(primitive2)
                    Case Token.Or
                        Return Primitive.op_Or(primitive, primitive2)
                    Case Token.Subtraction
                        Return primitive.Subtract(primitive2)
                    Case Else
                        Return Nothing
                End Select
            End Function

            Private Function EvaluateIdentifierExpression(ByVal expression As IdentifierExpression) As Primitive
                Dim value As Primitive = Nothing

                If Fields.TryGetValue(expression.Identifier.NormalizedText, value) Then
                    Return value
                End If

                Return Nothing
            End Function

            Private Function EvaluateLiteralExpression(ByVal expression As LiteralExpression) As Primitive
                If expression.Literal.Token = Token.StringLiteral Then
                    Return New Primitive(expression.Literal.Text.Trim(""""c))
                End If

                Return New Primitive(expression.Literal.Text)
            End Function

            Private Function EvaluateMethodCallExpression(ByVal expression As MethodCallExpression) As Primitive
                Dim typeInfo = TypeInfoBag.Types(expression.TypeName.NormalizedText)
                Dim methodInfo = typeInfo.Methods(expression.MethodName.NormalizedText)
                Dim list As List(Of Object) = New List(Of Object)()

                For Each argument In expression.Arguments
                    list.Add(EvaluateExpression(argument))
                Next

                Dim obj As Object = methodInfo.Invoke(Nothing, list.ToArray())

                If TypeOf obj Is Primitive Then
                    Return obj
                End If

                Return Nothing
            End Function

            Private Function EvaluateNegativeExpression(ByVal expression As NegativeExpression) As Primitive
                Return -EvaluateExpression(expression.Expression)
            End Function

            Private Function EvaluatePropertyExpression(ByVal expression As PropertyExpression) As Primitive
                Dim typeInfo = TypeInfoBag.Types(expression.TypeName.NormalizedText)
                Dim propertyInfo = typeInfo.Properties(expression.PropertyName.NormalizedText)
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
            Private Set(ByVal value As AppDomain)
                _DebuggerAppDomain = value
            End Set
        End Property

        Public Property Compiler As Compiler
            Get
                Return _Compiler
            End Get
            Private Set(ByVal value As Compiler)
                _Compiler = value
            End Set
        End Property

        Public Property CurrentDebuggerState As DebuggerState
            Get
                Return _CurrentDebuggerState
            End Get
            Private Set(ByVal value As DebuggerState)
                _CurrentDebuggerState = value
            End Set
        End Property

        Public Property CurrentInstruction As Instruction
            Get
                Return _CurrentInstruction
            End Get
            Private Set(ByVal value As Instruction)
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
            Private Set(ByVal value As ProgramTranslator)
                _Translator = value
            End Set
        End Property

        Private Shared Property Runner As ProgramRunner
        Public Event DebuggerStateChanged As EventHandler
        Public Event EngineUnloaded As EventHandler
        Public Event LineNumberChanged As EventHandler

        Public Sub New(ByVal compiler As Compiler)
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

        Public Sub AddBreakpoint(ByVal lineNumber As Integer)
            DebuggerAppDomain.DoCallBack(Sub() Runner.Breakpoints.Add(lineNumber))
        End Sub

        Public Sub [Continue]()
            DebuggerAppDomain.DoCallBack(Sub() Runner.Continue())
        End Sub

        Public Sub Pause()
            DebuggerAppDomain.DoCallBack(Sub() Runner.Pause())
        End Sub

        Private Sub RaiseDebuggerStateChangedEvent(ByVal currentState As DebuggerState)
            CurrentDebuggerState = currentState
            RaiseEvent DebuggerStateChanged(Me, EventArgs.Empty)
        End Sub

        Private Sub RaiseLineNumberChangedEvent(ByVal currentInstruction As Instruction)
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

        Public Sub RunProgram(ByVal stopOnFirstInstruction As Boolean)
            DebuggerAppDomain.DoCallBack(Sub() Runner.RunProgram(stopOnFirstInstruction:=True))
        End Sub
    End Class
End Namespace
