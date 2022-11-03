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
        Private _Document As TextDocument
        Private _IsDebuggerActive As Boolean
        Private _ProgramEngine As ProgramEngine
        Private breakpointsField As New List(Of Integer)()

        Public ReadOnly Property Breakpoints As List(Of Integer)
            Get
                If ProgramEngine IsNot Nothing Then
                    Return ProgramEngine.Breakpoints
                End If

                Return breakpointsField
            End Get
        End Property

        Public ReadOnly Property Dispatcher As Dispatcher
            Get
                Return Document.EditorControl.Dispatcher
            End Get
        End Property

        Public Property Document As TextDocument
            Get
                Return _Document
            End Get
            Private Set(value As TextDocument)
                _Document = value
            End Set
        End Property

        Public Property IsDebuggerActive As Boolean
            Get
                Return _IsDebuggerActive
            End Get
            Private Set(value As Boolean)
                _IsDebuggerActive = value
            End Set
        End Property

        Public Property ProgramEngine As ProgramEngine
            Get
                Return _ProgramEngine
            End Get
            Private Set(value As ProgramEngine)
                _ProgramEngine = value
            End Set
        End Property

        Private Property ReadOnlyRegion As IReadOnlyRegionHandle

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

        Private Sub New(document As TextDocument)
            Me.Document = document
            DocumentTracker.SetDocumentProperty(Me.Document, Me)
        End Sub

        Private Sub EnsureDebuggerIsActive(breakOnStart As Boolean)
            If Not IsDebuggerActive Then
                Document.Errors.Clear()
                Dim compiler = CompilerService.Compile(Document.Text, Document.Errors)

                If Document.Errors.Count = 0 Then
                    ProgramEngine = New ProgramEngine(compiler)
                    AddHandler ProgramEngine.EngineUnloaded, AddressOf OnDebuggerUnloaded
                    AddHandler ProgramEngine.DebuggerStateChanged, AddressOf OnDebuggerCurrentStateChanged
                    ProgramEngine.RunProgram(breakOnStart)
                    ProgramEngine.Breakpoints.AddRange(breakpointsField)
                    IsDebuggerActive = True
                    ReadOnlyRegion = TextBuffer.ReadOnlyRegionManager.CreateReadOnlyRegion(0, TextBuffer.CurrentSnapshot.Length, SpanTrackingMode.EdgeInclusive)
                End If
            End If
        End Sub

        Public Shared Function GetDebugger(document As TextDocument) As ProgramDebugger
            Dim programDebugger = DocumentTracker.GetDocumentProperty(Of ProgramDebugger)(document)

            If programDebugger Is Nothing Then
                programDebugger = New ProgramDebugger(document)
            End If

            Return programDebugger
        End Function

        Private Sub OnDebuggerCurrentStateChanged(sender As Object, e As EventArgs)
            Dim lineNumber = ProgramEngine.CurrentInstruction.LineNumber
            Dim currentState = ProgramEngine.CurrentDebuggerState
            Dispatcher.BeginInvoke(
                Sub()
                    Dim stMarkerProvider = StatementMarkerProvider.GetStatementMarkerProvider(TextView)
                    stMarkerProvider.ClearAllMarkers()

                    If currentState = DebuggerState.Paused Then
                        Dim lineFromLineNumber = Document.TextBuffer.CurrentSnapshot.GetLineFromLineNumber(lineNumber)
                        Dim span As TextSpan = New TextSpan(lineFromLineNumber.TextSnapshot, lineFromLineNumber.Start, lineFromLineNumber.Length, SpanTrackingMode.EdgeExclusive)
                        stMarkerProvider.AddStatementMarker(New StatementMarker(span, Colors.Gold))
                        TextView.ViewScroller.EnsureSpanVisible(New Span(lineFromLineNumber.Start, lineFromLineNumber.Length), 0.0, 0.0)
                    End If
                End Sub)
        End Sub

        Public Sub StepInto()
            If Not IsDebuggerActive Then
                EnsureDebuggerIsActive(breakOnStart:=True)
            Else
                ProgramEngine.StepInto()
            End If
        End Sub

        Public Sub StepOver()
            If Not IsDebuggerActive Then
                EnsureDebuggerIsActive(breakOnStart:=True)
            Else
                ProgramEngine.StepOver()
            End If
        End Sub

        Private Sub OnDebuggerUnloaded(sender As Object, e As EventArgs)
            RemoveHandler ProgramEngine.DebuggerStateChanged, AddressOf OnDebuggerCurrentStateChanged
            RemoveHandler ProgramEngine.EngineUnloaded, AddressOf OnDebuggerUnloaded
            IsDebuggerActive = False
            ProgramEngine = Nothing
            ReadOnlyRegion.Remove()
            Dispatcher.BeginInvoke(
                Sub()
                    Dim sMarkerProvider = StatementMarkerProvider.GetStatementMarkerProvider(TextView)
                    sMarkerProvider.ClearAllMarkers()
                End Sub)
        End Sub
    End Class
End Namespace
