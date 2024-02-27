Imports System.Threading
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.Engine

    <Serializable>
    Public Class ProgramRunner
        Private previousLineNumber As Integer = -1

        Public Property Breakpoints As New List(Of Integer)

        Public Property CurrentInstruction As Instruction

        Public Property DebuggerCommand As DebuggerCommand

        Public Property DebuggerExecution As New ManualResetEvent(initialState:=True)

        Public Property DebuggerState As DebuggerState

        Public Property Instructions As List(Of Instruction)

        Public Property LabelMap As New Dictionary(Of String, Integer)

        Public Property SubroutineInstructions As Dictionary(Of String, List(Of Instruction))
        Public Property Engine As ProgramEngine
        Public Property Fields As New Dictionary(Of String, Library.Primitive)

        Friend SymbolTable As SymbolTable
        Friend TypeInfoBag As TypeInfoBag

        Public Event LineNumberChanged As EventHandler
        Public Event DebuggerStateChanged As EventHandler

        Public Sub New(appDomain As AppDomain)
            _Engine = appDomain.GetData("ProgramEngine")
            TypeInfoBag = Compiler.GetTypeInfoBag()
            SymbolTable = Engine.Compiler.Parser.SymbolTable
        End Sub

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

            Dim thread As New Thread(
                Sub() Execute(Engine.Compiler.Parser.ParseTree)
             )
            thread.IsBackground = True
            thread.Start()
        End Sub

        Friend Sub ExecuteInstructions(instructions As List(Of Instruction))
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
                Dim text = CurrentInstruction.Execute(Me)
                PrepareDebuggerForNextInstruction()
                previousLineNumber = CurrentInstruction.LineNumber
                num = If(Not Equals(text, Nothing), LabelMap(text), num + 1)
            End While

            ChangeDebuggerState(DebuggerState.Finished)
        End Sub

        Friend Function Execute(statements As List(Of Statement)) As Statement
            If statements.Count = 0 Then Return Nothing
            Dim startLine = statements(0).StartToken.Line
            Dim endLine = statements.Last.StartToken.Line

            For i = 0 To statements.Count - 1
                Dim st = statements(i)
                If TypeOf st Is SubroutineStatement Then Continue For
                Dim result = st.Execute(Me)
                If TypeOf result Is ReturnStatement OrElse TypeOf result Is JumpLoopStatement Then
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
    End Class


End Namespace
