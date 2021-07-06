' To configure or remove Option's included in result, go to Options/Advanced Options...
Option Compare Text
Option Explicit On
Option Infer On
Option Strict On
Imports System.CodeDom.Compiler
Imports System.ComponentModel
Imports System.Diagnostics

Namespace Microsoft.SmallBasic.Library.Internal
    ''' <remarks/>
    <GeneratedCode("wsdl", "3.5.20706.1")>
    <DebuggerStepThrough>
    <DesignerCategory("code")>
    Public Class StatusCompletedEventArgs
        Inherits AsyncCompletedEventArgs
        Private results As Object()
        ''' <remarks/> 
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
