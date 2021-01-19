Imports System
Imports System.Collections.Generic
Imports System.IO
Imports Microsoft.Nautilus.Text
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.LanguageService
    Public Module CompilerService
        Private _compiler As Compiler

        Public ReadOnly Property DummyCompiler As Compiler
            Get

                If _compiler Is Nothing Then
                    _compiler = New Compiler()
                    _compiler.Initialize()
                End If

                Return _compiler
            End Get
        End Property

        Public Event CurrentCompletionItemChanged As EventHandler(Of CurrentCompletionItemChangedEventArgs)

        Public Function Compile(ByVal programText As String, ByVal outputFilePath As String, ByVal errors As ICollection(Of String)) As Boolean
            Try
                Dim compiler As Compiler = New Compiler()
                compiler.Initialize()
                Dim fileNameWithoutExtension = Path.GetFileNameWithoutExtension(outputFilePath)
                Dim directoryName = Path.GetDirectoryName(outputFilePath)
                Dim list As List(Of [Error]) = compiler.Build(New StringReader(programText), fileNameWithoutExtension, directoryName)

                For Each item In list
                    errors.Add($"{item.Line + 1},{item.Column + 1}: {item.Description}")
                Next

                Return errors.Count = 0
            Catch ex As Exception
                errors.Add(ex.Message)
                Return False
            End Try
        End Function

        Public Function Compile(ByVal programText As String, ByVal errors As ICollection(Of String)) As Compiler
            Try
                Dim compiler As Compiler = New Compiler()
                compiler.Initialize()
                Dim list As List(Of [Error]) = compiler.Compile(New StringReader(programText))

                For Each item In list
                    errors.Add($"{item.Line + 1},{item.Column + 1}: {item.Description}")
                Next

                Return compiler
            Catch ex As Exception
                errors.Add(ex.Message)
                Return Nothing
            End Try
        End Function

        Public Sub UpdateCurrentCompletionItem(ByVal completionItemWrapper As CompletionItemWrapper)
            RaiseEvent CurrentCompletionItemChanged(Nothing, New CurrentCompletionItemChangedEventArgs(completionItemWrapper))
        End Sub

        Public Sub FormatDocument(ByVal textBuffer As ITextBuffer)
            Dim num = 2
            Dim currentSnapshot = textBuffer.CurrentSnapshot
            Dim source As TextBufferReader = New TextBufferReader(currentSnapshot)
            Dim completionHelper As CompletionHelper = New CompletionHelper()
            completionHelper.GetIndentationLevel(source, 0)
            Dim textEdit As ITextEdit = textBuffer.CreateEdit()

            Using textEdit

                For Each line In currentSnapshot.Lines
                    Dim positionOfNextNonWhiteSpaceCharacter = line.GetPositionOfNextNonWhiteSpaceCharacter(0)
                    Dim indentationLevel = completionHelper.GetIndentationLevel(line.LineNumber)

                    If positionOfNextNonWhiteSpaceCharacter <> indentationLevel * num Then
                        textEdit.Replace(line.Start, positionOfNextNonWhiteSpaceCharacter, New String(" "c, indentationLevel * num))
                    End If
                Next

                textEdit.Apply()
            End Using
        End Sub
    End Module
End Namespace
