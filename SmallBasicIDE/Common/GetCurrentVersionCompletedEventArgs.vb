Imports System
Imports System.CodeDom.Compiler
Imports System.ComponentModel
Imports System.Diagnostics

Namespace Microsoft.SmallBasic.com.smallbasic
    <DesignerCategory("code")>
    <GeneratedCode("System.Web.Services", "4.0.30319.1")>
    <DebuggerStepThrough>
    Public Class GetCurrentVersionCompletedEventArgs
        Inherits AsyncCompletedEventArgs

        Private results As Object()

        Public ReadOnly Property Result As String
            Get
                RaiseExceptionIfNecessary()
                Return CStr(results(0))
            End Get
        End Property

        Friend Sub New(ByVal results As Object(), ByVal exception As Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
    End Class
End Namespace
