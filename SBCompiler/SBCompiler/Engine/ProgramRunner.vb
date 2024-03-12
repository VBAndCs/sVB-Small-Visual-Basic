Imports System.Threading
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.Engine

    Public Class ProgramRunner
        Friend PreviousLineNumber As Integer = -1
        Friend CurrentLineNumber As Integer

        Public Property Breakpoints As New List(Of Integer)

        Public Property CurrentStatement As Statement

        Public Property DebuggerCommand As DebuggerCommand

        Public Property DebuggerState As DebuggerState

        Public Property Instructions As List(Of Instruction)

        Public Property LabelMap As New Dictionary(Of String, Integer)

        Public Property SubroutineInstructions As Dictionary(Of String, List(Of Instruction))

        Public DocLineOffset As Integer

        Public Property Engine As ProgramEngine
        Public Property Fields As New Dictionary(Of String, Library.Primitive)

        Friend SymbolTable As SymbolTable
        Friend TypeInfoBag As TypeInfoBag
        Public currentParser As Parser


        Public Event DebuggerStateChanged As EventHandler

        Public Sub New(engine As ProgramEngine, parser As Parser)
            DocLineOffset = parser.DocStartLine
            _Engine = engine
            TypeInfoBag = Compiler.TypeInfoBag
            currentParser = parser
            SymbolTable = currentParser.SymbolTable
            _Engine.runners(currentParser) = Me
            _Engine.CurrentDebuggerState = _DebuggerState
            RaiseEvent DebuggerStateChanged(Me, EventArgs.Empty)
        End Sub

        Dim isGlobalInitialized As Boolean = False

        Friend ReadOnly Property GlobalRunner As ProgramRunner
            Get
                Dim gRunner = _Engine.runners(_Engine.GlobalParser)
                If Not isGlobalInitialized Then
                    isGlobalInitialized = True
                    gRunner.Execute(_Engine.GlobalParser.ParseTree)
                End If
                Return gRunner
            End Get
        End Property

        Friend Sub RunForm(formName As String)
            Dim className = "_smallvisualbasic_" & formName.ToLower()
            Dim formParser As Parser
            For Each p In _Engine.Parsers
                If p.ClassName.ToLower() = className Then
                    formParser = p
                    Exit For
                End If
            Next

            If formParser Is Nothing Then
                Throw New Exception("There is no form name {formName} in the project")
            End If

            Dim formRunner As ProgramRunner
            If _Engine.runners.ContainsKey(formParser) Then
                formRunner = _Engine.runners(formParser)
            Else
                formRunner = New ProgramRunner(_Engine, formParser)
            End If
            formRunner.Execute(formParser.ParseTree)
        End Sub

        Friend Sub SetGlobalField(name As String, value As Library.Primitive)
            GlobalRunner.Fields(name) = value
        End Sub

        Friend Function GetGlobalField(name As String) As Library.Primitive
            If GlobalRunner.Fields.ContainsKey(name) Then
                Return GlobalRunner.Fields(name)
            Else
                Return 0
            End If
        End Function

        Friend Function GetKey(identifier As Token) As String
            Dim variableName = identifier.LCaseText
            Dim Subroutine = identifier.SubroutineName?.ToLower()
            If Subroutine = "" Then Return variableName

            Dim key = $"{Subroutine}.{variableName}"
            If SymbolTable.LocalVariables.ContainsKey(key) Then
                Return key
            Else
                Return variableName
            End If
        End Function

        Private Sub ChangeDebuggerState(state As DebuggerState)
            If _DebuggerState <> state Then
                _DebuggerState = state
                _Engine.CurrentDebuggerState = state
                RaiseEvent DebuggerStateChanged(Me, EventArgs.Empty)
            End If
        End Sub

        Friend Sub CheckForExecutionBreakAtLine(lineNumber As Integer, Optional isSub As Boolean = False)
            If DoStepOver Then
                If Depth > 0 OrElse (isSub AndAlso lineNumber <> StepOverLineNumber) Then
                    Return
                End If
            End If

            CurrentLineNumber = lineNumber
            CheckForExecutionBreak(False)
            PreviousLineNumber = lineNumber
        End Sub

        Private Sub CheckForExecutionBreak(isFirstStatement As Boolean)
            If isFirstStatement AndAlso _Engine.StopOnFirstStaement Then
                _Engine.StopOnFirstStaement = False
                Pause()
                ChangeDebuggerState(DebuggerState.Running)

            ElseIf CurrentLineNumber <> PreviousLineNumber Then
                If Breakpoints.Contains(CurrentLineNumber - DocLineOffset) OrElse
                        DebuggerCommand = DebuggerCommand.StepInto OrElse
                        DebuggerCommand = DebuggerCommand.StepOver Then
                    Pause()
                    ChangeDebuggerState(DebuggerState.Running)
                End If
            End If
        End Sub


        Friend CurrentThread As Thread

        Public Sub [Continue]()
            isPaused = False
            If CurrentThread.IsAlive Then
                Try
                    CurrentThread.Resume()
                Catch
                End Try
            End If
        End Sub

        Dim isPaused As Boolean

        Public Sub Pause()
            isPaused = True
            ChangeDebuggerState(DebuggerState.Paused)
            CurrentThread.Suspend()
        End Sub

        Public Sub Reset()
            Fields.Clear()
            Depth = 0
        End Sub


        Friend Depth As Integer = 0
        Friend DoStepOver As Boolean = False
        Friend StepOverLineNumber As Integer = -1

        Public Sub StepInto()
            DebuggerCommand = DebuggerCommand.StepInto
            Depth = 0
            DoStepOver = False
            [Continue]()
        End Sub

        Public Sub StepOver()
            DebuggerCommand = DebuggerCommand.StepOver
            Depth = 0
            DoStepOver = True
            StepOverLineNumber = CurrentLineNumber
            [Continue]()
        End Sub


        Friend Function Execute(statements As List(Of Statement)) As Statement
            PreviousLineNumber = -1
            If statements.Count = 0 Then Return Nothing

            Dim startLine = statements(0).StartToken.Line
            Dim endLine = statements.Last.StartToken.Line
            _Engine.CurrentRunner = Me
            Dim isFirstStatement = True

            For i = 0 To statements.Count - 1
                Dim st = statements(i)
                If TypeOf st Is SubroutineStatement OrElse
                    TypeOf st Is EmptyStatement OrElse
                    TypeOf st Is IllegalStatement Then Continue For

                CurrentStatement = st
                CurrentLineNumber = CurrentStatement.StartToken.Line
                Dim isNotGenCode = CurrentLineNumber - DocLineOffset > -1

                If isNotGenCode Then
                    If Not DoStepOver OrElse Depth = 0 Then
                        CheckForExecutionBreak(isFirstStatement)
                    End If
                End If

                PreviousLineNumber = CurrentLineNumber
                Dim result = st.Execute(Me)

                If isNotGenCode Then isFirstStatement = False
                If TypeOf result Is EndDebugging Then Return result

                If SmallVisualBasic.Library.Program.IsTerminated Then
                    Return New EndDebugging()
                ElseIf TypeOf result Is ReturnStatement OrElse TypeOf result Is JumpLoopStatement Then
                    Return result
                ElseIf TypeOf result Is GotoStatement Then
                    Dim gotoLine = CType(result, GotoStatement).Label.Line
                    If gotoLine < startLine OrElse gotoLine > endLine Then
                        Return result
                    Else
                        For j = 0 To statements.Count - 1
                            If statements(j).StartToken.Line = gotoLine Then
                                i = j - 1
                                Exit For
                            End If
                        Next
                    End If
                End If
            Next

            Return Nothing
        End Function

        Friend Sub RaiseDebuggerStateChanged()
            RaiseEvent DebuggerStateChanged(_Engine.CurrentRunner, EventArgs.Empty)
        End Sub
    End Class


End Namespace
