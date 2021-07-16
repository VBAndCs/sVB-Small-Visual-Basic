Imports System
Imports System.IO

Namespace Microsoft.SmallBasic
    Friend NotInheritable Class Program
        Private _text3 As String = VisualBasic.Constants.vbCrLf & "i = 15" & VisualBasic.Constants.vbCrLf & "j = 23" & VisualBasic.Constants.vbCrLf & "if (j >= i AND j <= i * 20) then" & VisualBasic.Constants.vbCrLf & "TextWindow.WriteLine(""Foo"")" & VisualBasic.Constants.vbCrLf & "endif" & VisualBasic.Constants.vbCrLf & "TextWindow.Pause()" & VisualBasic.Constants.vbCrLf & VisualBasic.Constants.vbCrLf & "Sub MySub" & VisualBasic.Constants.vbCrLf & "  TextWindow.WriteLine(i)" & VisualBasic.Constants.vbCrLf & "EndSub" & VisualBasic.Constants.vbCrLf

        Public Shared Sub Main(ByVal args As String())
            Dim program As Program = New Program()

            If args.Length <> 1 Then
                Console.WriteLine("Usage: SmallBasicCompiler.exe <SB File>")
            Else
                program.Run(args(0))
            End If
        End Sub

        Private Sub Run(ByVal file As String)
            If Not IO.File.Exists(file) Then
                Console.WriteLine("{0} doesn't exist.", file)
                Return
            End If

            Using source As StreamReader = New StreamReader(file)
                Dim compiler As Compiler = New Compiler()
                Path.GetDirectoryName(file)
                Dim fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file)
                compiler.Build(source, fileNameWithoutExtension, Directory.GetCurrentDirectory())
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

            Dim compiler As Compiler = New Compiler()
            compiler.Build(New StringReader(_text3), "foo", "c:\users\vijayeg\documents\smallbasic\")
        End Sub

        Private Sub TestBuild()
            Dim files = Directory.GetFiles("C:\Users\vijayeg\Documents\smallbasic\samples", "*.smallbasic")

            For Each _text In files

                Using source As StreamReader = New StreamReader(_text)
                    Dim compiler As Compiler = New Compiler()
                    Dim directoryName = Path.GetDirectoryName(_text)
                    Dim fileNameWithoutExtension = Path.GetFileNameWithoutExtension(_text)
                    compiler.Build(source, fileNameWithoutExtension, directoryName)
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

                Using reader As StreamReader = New StreamReader(_text)
                    Dim parser As Parser = New Parser()
                    parser.Parse(reader)
                    Console.Write(_text)
                    Console.WriteLine(": {0} errors.", parser.Errors.Count)

                    If parser.SymbolTable.Variables.Count > 0 Then
                        Console.WriteLine("** Variables **")

                        For Each value In parser.SymbolTable.Variables.Values
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
            Dim parser As Parser = New Parser()

            While True
                Console.Write("> ")
                Dim lineText As String = Console.ReadLine()
                Dim lineScanner As LineScanner = New LineScanner()
                Dim tokenList = lineScanner.GetTokenList(lineText, 0)
                Dim expression = parser.BuildLogicalExpression(tokenList)
                Console.WriteLine("= {0}", Parser.EvaluateExpression(expression))
                Console.WriteLine(expression)
                Console.WriteLine()
            End While
        End Sub
    End Class
End Namespace
