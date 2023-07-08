Imports System
Imports System.IO

Namespace Microsoft.SmallVisualBasic
    Friend NotInheritable Class Program

        Private _text3 As String = "
i = 15
j = 23
if (j >= i AND j <= i * 20) then
    TextWindow.WriteLine(""Foo"")
endif
TextWindow.Pause()

Sub MySub
    TextWindow.WriteLine(i)
EndSub"

        Public Shared Sub Main(args As String())
            Dim program As New Program()

            If args.Length <> 1 Then
                Console.WriteLine("Usage: SmallBasicCompiler.exe <SB File>")
            Else
                program.Run(args(0))
            End If
        End Sub

        Private Sub Run(file As String)
            If Not IO.File.Exists(file) Then
                Console.WriteLine("{0} doesn't exist.", file)
                Return
            End If

            Using source As New StreamReader(file)
                Dim compiler As New Compiler()
                Path.GetDirectoryName(file)
                Dim fileName = Path.GetFileNameWithoutExtension(file)
                compiler.Compile(source)
                Dim parsers As New List(Of Parser)
                parsers.Add(compiler.Parser)
                compiler.Build(parsers, fileName, Directory.GetCurrentDirectory())
                Console.Write(file)
                Console.WriteLine(": {0} errors.", compiler.Parser.Errors.Count)
                Console.WriteLine()
            End Using
        End Sub

        Private Sub TestParseTree()
            Dim parser As New Parser()
            Dim reader As New StringReader(_text3)
            parser.Parse(reader)

            For Each item In parser.ParseTree
                Console.Write(item)
            Next

            Dim compiler As New Compiler()
            compiler.Compile(New StringReader(_text3))
            Dim parsers As New List(Of Parser)
            parsers.Add(compiler.Parser)
            compiler.Build(parsers, "foo", "c:\users\vijayeg\documents\smallbasic\")
        End Sub

        Private Sub TestBuild()
            Dim files = Directory.GetFiles("C:\Users\vijayeg\Documents\smallbasic\samples", "*.smallbasic")

            For Each _text In files
                Using source As New StreamReader(_text)
                    Dim compiler As New Compiler()
                    Dim directoryName = Path.GetDirectoryName(_text)
                    Dim fileNameWithoutExtension = Path.GetFileNameWithoutExtension(_text)
                    compiler.Compile(source)
                    Dim parsers As New List(Of Parser)
                    parsers.Add(compiler.Parser)
                    compiler.Build(parsers, fileNameWithoutExtension, directoryName)
                    Console.Write(_text)
                    Console.WriteLine(": {0} errors.", compiler.Parser.Errors.Count)
                    Console.WriteLine()
                    Path.Combine(directoryName, fileNameWithoutExtension & ".exe")
                End Using
            Next
        End Sub

        Private Sub TestParseTree2()
            Dim files = Directory.GetFiles("C:\Users\vijayeg\Documents\smallbasic\samples", "*.smallbasic")

            For Each _text In files
                Using reader As New StreamReader(_text)
                    Dim parser As New Parser()
                    parser.Parse(reader)
                    Console.Write(_text)
                    Console.WriteLine(": {0} errors.", parser.Errors.Count)

                    If parser.SymbolTable.GlobalVariables.Count > 0 Then
                        Console.WriteLine("** Variables **")

                        For Each value In parser.SymbolTable.GlobalVariables.Values
                            Console.Write(value.Text & "; ")
                        Next

                        Console.WriteLine()
                    End If

                    If parser.SymbolTable.Labels.Count > 0 Then
                        Console.WriteLine("** Labels **")

                        For Each value2 In parser.SymbolTable.Labels.Values
                            Console.Write(value2.Text & "; ")
                        Next

                        Console.WriteLine()
                    End If

                    If parser.SymbolTable.Subroutines.Count > 0 Then
                        Console.WriteLine("** Subroutines **")

                        For Each value3 In parser.SymbolTable.Subroutines.Values
                            Console.Write(value3.Text & "; ")
                        Next

                        Console.WriteLine()
                    End If

                    Console.WriteLine()
                End Using
            Next
        End Sub

        Private Sub TestExpression()
            Dim parser As New Parser()

            While True
                Console.Write("> ")
                Dim lineText = Console.ReadLine()
                Dim tokenList = LineScanner.GetTokenEnumerator(lineText, 0)
                Dim expression = parser.BuildLogicalExpression(tokenList)
                Console.WriteLine("= {0}", Parser.EvaluateExpression(expression))
                Console.WriteLine(expression)
                Console.WriteLine()
            End While
        End Sub
    End Class
End Namespace
