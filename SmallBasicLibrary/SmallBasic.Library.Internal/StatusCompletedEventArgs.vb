Imports System.CodeDom.Compiler
Imports System.ComponentModel

Namespace Library.Internal

    <GeneratedCode("wsdl", "3.5.20706.1")>
    <DebuggerStepThrough>
    <DesignerCategory("code")>
    Public Class StatusCompletedEventArgs
        Inherits AsyncCompletedEventArgs
        Private results As Object()

        Public ReadOnly Property Result As String
            Get
                RaiseExceptionIfNecessary()
                Return CStr(results(0))
            End Get
        End Property

        Friend Sub New(results As Object(), exception As Exception, cancelled As Boolean, userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
    End Class
End Namespace
