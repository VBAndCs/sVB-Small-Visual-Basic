
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' The Program class provides helpers to control the program execution.
    ''' </summary>
    <SmallVisualBasicType>
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
                Dim asm = Assembly.GetCallingAssembly()
                Return Path.GetDirectoryName(asm.Location)
            End Get
        End Property

        ''' <summary>
        ''' Delays program execution by the specified amount of milliseconds.
        ''' If this method caused any troubles, please use the WinDelay method
        ''' </summary>
        ''' <param name="milliSeconds">The amount of delay.</param>
        Public Shared Sub Delay(milliSeconds As Primitive)
            Dim t = CDbl(milliSeconds)
            If t <= 0 Then Return
            Thread.Sleep(TimeSpan.FromMilliseconds(t))
        End Sub

        ''' <summary>
        ''' Delays program execution by the specified amount of milliseconds.
        ''' If this method caused any troubles, please use the Delay method
        ''' </summary>
        ''' <param name="milliSeconds">The amount of delay.</param>
        Public Shared Sub WinDelay(milliSeconds As Primitive)
            Dim t = CDbl(milliSeconds)
            If t <= 0 Then Return

            SmallBasicApplication.Invoke(
                Sub()
                    Thread.Sleep(TimeSpan.FromMilliseconds(t))
                End Sub)
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
        ''' <param name="index">Index of the argument.</param>
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

        ''' <summary>
        ''' Gets a value from the regestry.
        ''' </summary>
        ''' <param name="section">The regetry section name that you used when saving the setting.</param>
        ''' <param name="name">The regestry key that you used when saving the setting.</param>
        ''' <param name="defaultValue">The value to return if the data is not found in the regestry.</param>
        ''' <returns>A string that contains the required setting.</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetSetting(section As Primitive, name As Primitive, defaultValue As Primitive) As Primitive
            Dim appName = "sVB_" & IO.Path.GetDirectoryName(Directory).Replace(" ", "_")
            Return Interaction.GetSetting(appName, section, name, defaultValue)
        End Function

        ''' <summary>
        ''' Saves a value in the regestry.
        ''' </summary>
        ''' <param name="section">The regetry section name, which refers to the category that the data lies under. For example, you may use "Form1" as a section name when you save property values of this form.</param>
        ''' <param name="name">The regestry key where the data is saved in. Shows a key name that refers to the meaning of the date, like "Width" and "Height" if you are saving the width and the height of the form.</param>
        ''' <param name="value">The value to save in the regestry.</param>
        Public Shared Sub SaveSetting(section As Primitive, name As Primitive, value As Primitive)
            Dim appName = "sVB_" & IO.Path.GetDirectoryName(Directory).Replace(" ", "_")
            Interaction.SaveSetting(appName, section, name, value)
        End Sub
    End Class
End Namespace
