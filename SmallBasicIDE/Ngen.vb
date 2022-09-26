Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows

Namespace Microsoft.SmallVisualBasic
    <RunInstaller(True)>
    Public Class Ngen
        Inherits Installer

        Private components As IContainer

        Public Sub New()
            InitializeComponent()
        End Sub

        Public Overrides Sub Install(stateSaver As IDictionary)
            MyBase.Install(stateSaver)

            Try
                Dim fileName As String = Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "ngen.exe")
                Dim process As Process = New Process()
                process.StartInfo.FileName = fileName
                Dim arg = Context.Parameters("assemblypath")
                process.StartInfo.CreateNoWindow = True
                process.StartInfo.RedirectStandardOutput = True
                process.StartInfo.UseShellExecute = False
                process.StartInfo.Arguments = $"install ""{arg}"" /silent"
                process.Start()
                process.StandardOutput.ReadToEnd()
                process.WaitForExit()
            Catch ex As Exception
                MessageBox.Show(ex.ToString(), "Small Basic", MessageBoxButton.OK, MessageBoxImage.Hand)
            End Try
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

        Private Sub InitializeComponent()
            components = New Container()
        End Sub
    End Class
End Namespace
