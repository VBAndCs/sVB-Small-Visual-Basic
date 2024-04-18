Imports System.IO

Namespace Microsoft.SmallVisualBasic
    Friend NotInheritable Class Program

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
    End Class
End Namespace
