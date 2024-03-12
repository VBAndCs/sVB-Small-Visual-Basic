Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.SmallVisualBasic.Engine
Imports Microsoft.SmallVisualBasic.LanguageService
Imports Microsoft.Windows.Controls
Imports System
Imports System.Collections.Generic
Imports System.Windows.Media
Imports System.Windows.Threading

Namespace Microsoft.SmallVisualBasic.Documents
    Public Class ProgramDebugger

        Public ReadOnly Property Dispatcher As Dispatcher
            Get
                Return Document.EditorControl.Dispatcher
            End Get
        End Property

        Public ReadOnly Property Document As TextDocument
            Get
                Dim docPath = ProgramEngine.CurrentRunner.currentParser.DocPath
                Dim m = Helper.MainWindow
                Dim doc = If(docPath = "", m.ActiveDocument, m.OpenDocIfNot(docPath))
                Return doc
            End Get
        End Property

        Public Property IsDebuggerActive As Boolean

        Public Property ProgramEngine As ProgramEngine

        Private Function CreateEngine() As ProgramEngine
            If _ProgramEngine Is Nothing Then
                Dim parsers = Helper.MainWindow.BuildAndRun(True)
                If parsers Is Nothing OrElse parsers.Count = 0 Then Return Nothing

                _ProgramEngine = New Engine.ProgramEngine(parsers)
                AddHandler Library.Program.ProgramTerminated, AddressOf OnProgramTerminated
                AddHandler _ProgramEngine.CurrentRunner.DebuggerStateChanged, AddressOf OnDebuggerCurrentStateChanged
            End If
            Return _ProgramEngine
        End Function


        Public ReadOnly Property TextBuffer As ITextBuffer
            Get
                Return Document.TextBuffer
            End Get
        End Property

        Public ReadOnly Property TextView As ITextView
            Get
                Return Document.EditorControl.TextView
            End Get
        End Property

        Dim _project As String

        Private Sub New(project As String)
            _project = project.ToLower()
        End Sub


        Private Shared debuggers As New Dictionary(Of String, ProgramDebugger)

        Public Shared Function GetDebugger(project As String) As ProgramDebugger
            Dim key = project.ToLower()
            If debuggers.ContainsKey(key) Then Return debuggers(key)

            Dim debugger As New ProgramDebugger(project)
            debuggers(key) = debugger
            Return debugger
        End Function

        Public Sub Run(breakOnStart As Boolean)
            If IsDebuggerActive Then
                _ProgramEngine.Continue()
            Else
                _ProgramEngine = CreateEngine()
                If _ProgramEngine IsNot Nothing Then
                    _ProgramEngine.RunProgram(breakOnStart)
                    _ProgramEngine.CurrentRunner.Breakpoints = Document.Breakpoints
                    _IsDebuggerActive = True
                    Document.ReadOnlyRegion = TextBuffer.ReadOnlyRegionManager.CreateReadOnlyRegion(0, TextBuffer.CurrentSnapshot.Length, SpanTrackingMode.EdgeInclusive)
                End If
            End If
        End Sub

        Private Sub OnDebuggerCurrentStateChanged(sender As Object, e As EventArgs)
            Dim offset = _ProgramEngine.CurrentRunner.DocLineOffset
            Dim lineNumber = _ProgramEngine.CurrentLineNumber - offset
            Dim currentState = _ProgramEngine.CurrentDebuggerState

            Helper.MainWindow.Dispatcher.BeginInvoke(
                Sub()
                    Dim doc = Document
                    Dim tv = TextView
                    Dim stMarkerProvider = StatementMarkerProvider.GetStatementMarkerProvider(tv)
                    Dim marker = doc.ExecutionMarker

                    If marker IsNot Nothing Then
                        doc.ExecutionMarker = Nothing
                        If doc.Breakpoints.Contains(marker.LineNumber) Then
                            marker.MarkerColor = Colors.Red
                        Else
                            stMarkerProvider.RemoveMarker(marker)
                        End If
                    End If

                    If currentState = DebuggerState.Paused Then
                        Dim line = doc.TextBuffer.CurrentSnapshot.GetLineFromLineNumber(lineNumber)

                        If doc.Breakpoints.Contains(lineNumber) Then
                            marker = stMarkerProvider.GetMarker(lineNumber)
                            marker.MarkerColor = Colors.DarkGoldenrod
                        Else
                            Dim span = doc.GetFullStatementSpan(lineNumber)
                            marker = New StatementMarker(span, lineNumber, Colors.Gold)
                            stMarkerProvider.AddStatementMarker(marker)
                        End If

                        doc.ExecutionMarker = marker
                        doc.EnsureLineVisible(line)
                    End If
                End Sub)


            ActivateDoc()
        End Sub

        Public Sub ActivateDoc()
            Dim m = Helper.MainWindow
            m.Dispatcher.BeginInvoke(
                  Sub()
                      If m.WindowState = System.Windows.WindowState.Minimized Then
                          m.WindowState = System.Windows.WindowState.Maximized
                      End If
                      m.Activate()
                      m.Focus()
                      m.tabCode.IsSelected = True
                  End Sub)
        End Sub

        Public Sub StepInto()
            If _IsDebuggerActive Then
                _ProgramEngine.StepInto()
            Else
                Run(breakOnStart:=True)
            End If
        End Sub

        Public Sub [Continue]()
            _ProgramEngine.Continue()
        End Sub

        Public Sub StepOver()
            If _IsDebuggerActive Then
                _ProgramEngine.StepOver()
            Else
                Run(breakOnStart:=True)
            End If
        End Sub

        Private Sub OnProgramTerminated()
            RemoveHandler _ProgramEngine.CurrentRunner.DebuggerStateChanged, AddressOf OnDebuggerCurrentStateChanged
            RemoveHandler Library.Program.ProgramTerminated, AddressOf OnProgramTerminated

            _IsDebuggerActive = False
            _ProgramEngine.Disopese()
            _ProgramEngine = Nothing

            For Each view As Shell.MdiView In Helper.MainWindow.viewsControl.Items
                Dim doc = view.Document
                doc.ReadOnlyRegion?.Remove()
                doc.EditorControl.Dispatcher.BeginInvoke(
                    Sub()
                        Dim sMarkerProvider = StatementMarkerProvider.GetStatementMarkerProvider(doc.EditorControl.TextView)
                        sMarkerProvider.ClearAllMarkers()
                    End Sub)
            Next
        End Sub
    End Class
End Namespace
