
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' The Program class provides helpers to control the program execution.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Program
        Private Shared args As String() = Environment.GetCommandLineArgs()

        ''' <summary>
        ''' Gets the number of command-line arguments passed to this program.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared ReadOnly Property ArgumentCount As Primitive
            Get
                Return If((args.Length <> 0), (args.Length - 1), 0)
            End Get
        End Property

        ''' <summary>
        ''' Gets the executing program's directory.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property Directory As Primitive
            Get
                Dim entryAssembly As Assembly = Assembly.GetEntryAssembly()
                Return Path.GetDirectoryName(entryAssembly.Location)
            End Get
        End Property

        ''' <summary>
        ''' Delays program execution by the specified amount of MilliSeconds.
        ''' </summary>
        ''' <param name="milliSeconds">
        ''' The amount of delay.
        ''' </param>
        Public Shared Sub Delay(milliSeconds As Primitive)
            Thread.Sleep(TimeSpan.FromMilliseconds(milliSeconds))
        End Sub

        ''' <summary>
        ''' Ends the program.
        ''' </summary>
        Public Shared Sub [End]()
            SmallBasicApplication.End()
        End Sub

        ''' <summary>
        ''' Returns the specified argument passed to this program.
        ''' </summary>
        ''' <param name="index">
        ''' Index of the argument.
        ''' </param>
        ''' <returns>
        ''' The command-line argument at the specified index.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetArgument(index As Primitive) As Primitive
            Dim num As Integer = index
            If num >= 1 AndAlso num < args.Length Then
                Return New Primitive(args(num))
            End If

            Return New Primitive("")
        End Function
    End Class
End Namespace
