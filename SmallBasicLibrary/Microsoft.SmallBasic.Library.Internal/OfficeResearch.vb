' To configure or remove Option's included in result, go to Options/Advanced Options...
Option Compare Text
Option Explicit On
Option Infer On
Option Strict On
Imports System.CodeDom.Compiler
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Threading
Imports System.Web.Services
Imports System.Web.Services.Description
Imports System.Web.Services.Protocols

Namespace Microsoft.SmallBasic.Library.Internal
    ''' <remarks/>
    <GeneratedCode("wsdl", "3.5.20706.1")>
    <DebuggerStepThrough>
    <DesignerCategory("code")>
    <WebServiceBinding(Name:="ResearchSoap", [Namespace]:="urn:Microsoft.Search")>
    Public Class OfficeResearch
        Inherits SoapHttpClientProtocol
        Private QueryOperationCompleted As SendOrPostCallback
        Private RegistrationOperationCompleted As SendOrPostCallback
        Private StatusOperationCompleted As SendOrPostCallback
        Private DiscoveryOperationCompleted As SendOrPostCallback

        ''' <remarks/>
        Public Event QueryCompleted As QueryCompletedEventHandler

        ''' <remarks/>
        Public Event RegistrationCompleted As RegistrationCompletedEventHandler

        ''' <remarks/>
        Public Event StatusCompleted As StatusCompletedEventHandler

        ''' <remarks/>
        Public Event DiscoveryCompleted As DiscoveryCompletedEventHandler

        ''' <remarks/>
        Public Sub New()
            MyBase.Url = "http://rr.office.microsoft.com/Research/query.asmx"
        End Sub
        ''' <remarks/>
        <SoapDocumentMethod("urn:Microsoft.Search/Query", RequestNamespace:="urn:Microsoft.Search", ResponseNamespace:="urn:Microsoft.Search", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function Query(queryXml As String) As String
            Dim array As Object() = Invoke("Query", New Object(0) {
                queryXml
            })
            Return CStr(array(0))
        End Function
        ''' <remarks/>
        Public Function BeginQuery(queryXml As String, callback As AsyncCallback, asyncState As Object) As IAsyncResult
            Return BeginInvoke("Query", New Object(0) {
                queryXml
            }, callback, asyncState)
        End Function
        ''' <remarks/>
        Public Function EndQuery(asyncResult As IAsyncResult) As String
            Dim array As Object() = EndInvoke(asyncResult)
            Return CStr(array(0))
        End Function
        ''' <remarks/>
        Public Sub QueryAsync(queryXml As String)
            QueryAsync(queryXml, Nothing)
        End Sub
        ''' <remarks/>
        Public Sub QueryAsync(queryXml As String, userState As Object)
            If QueryOperationCompleted Is Nothing Then
                QueryOperationCompleted = AddressOf OnQueryOperationCompleted
            End If
            InvokeAsync("Query", New Object(0) {
                queryXml
            }, QueryOperationCompleted, userState)
        End Sub

        Private Sub OnQueryOperationCompleted(arg As Object)
            Dim invokeCompletedEventArgs1 As InvokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
            RaiseEvent QueryCompleted(Me, New QueryCompletedEventArgs(invokeCompletedEventArgs1.Results, invokeCompletedEventArgs1.[Error], invokeCompletedEventArgs1.Cancelled, invokeCompletedEventArgs1.UserState))
        End Sub

        ''' <remarks/>
        <SoapDocumentMethod("urn:Microsoft.Search/Registration", RequestNamespace:="urn:Microsoft.Search", ResponseNamespace:="urn:Microsoft.Search", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function Registration(registrationXml As String) As String
            Dim array As Object() = Invoke("Registration", New Object(0) {registrationXml})
            Return CStr(array(0))
        End Function

        ''' <remarks/>
        Public Function BeginRegistration(registrationXml As String, callback As AsyncCallback, asyncState As Object) As IAsyncResult
            Return BeginInvoke("Registration", New Object(0) {
                registrationXml
            }, callback, asyncState)
        End Function

        Public Function EndRegistration(asyncResult As IAsyncResult) As String
            Dim array As Object() = EndInvoke(asyncResult)
            Return CStr(array(0))
        End Function


        Public Sub RegistrationAsync(registrationXml As String)
            RegistrationAsync(registrationXml, Nothing)
        End Sub


        Public Sub RegistrationAsync(registrationXml As String, userState As Object)
            If RegistrationOperationCompleted Is Nothing Then
                RegistrationOperationCompleted = AddressOf OnRegistrationOperationCompleted
            End If
            InvokeAsync("Registration", New Object(0) {
                registrationXml
            }, RegistrationOperationCompleted, userState)
        End Sub
        Private Sub OnRegistrationOperationCompleted(arg As Object)
            Dim invokeCompletedEventArgs1 As InvokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
            RaiseEvent RegistrationCompleted(Me, New RegistrationCompletedEventArgs(invokeCompletedEventArgs1.Results, invokeCompletedEventArgs1.[Error], invokeCompletedEventArgs1.Cancelled, invokeCompletedEventArgs1.UserState))
        End Sub


        <SoapDocumentMethod("urn:Microsoft.Search/Status", RequestNamespace:="urn:Microsoft.Search", ResponseNamespace:="urn:Microsoft.Search", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function Status() As String
            Dim array As Object() = Invoke("Status", New Object() {})
            Return CStr(array(0))
        End Function

        Public Function BeginStatus(callback As AsyncCallback, asyncState As Object) As IAsyncResult
            Return BeginInvoke("Status", New Object() {}, callback, asyncState)
        End Function

        Public Function EndStatus(asyncResult As IAsyncResult) As String
            Dim array As Object() = EndInvoke(asyncResult)
            Return CStr(array(0))
        End Function

        Public Sub StatusAsync()
            StatusAsync(Nothing)
        End Sub

        Public Sub StatusAsync(userState As Object)
            If StatusOperationCompleted Is Nothing Then
                StatusOperationCompleted = AddressOf OnStatusOperationCompleted
            End If
            InvokeAsync("Status", New Object() {}, StatusOperationCompleted, userState)
        End Sub

        Private Sub OnStatusOperationCompleted(arg As Object)
            Dim invokeCompletedEventArgs1 As InvokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
            RaiseEvent StatusCompleted(Me, New StatusCompletedEventArgs(invokeCompletedEventArgs1.Results, invokeCompletedEventArgs1.[Error], invokeCompletedEventArgs1.Cancelled, invokeCompletedEventArgs1.UserState))
        End Sub


        <SoapDocumentMethod("urn:Microsoft.Search/Discovery", RequestNamespace:="urn:Microsoft.Search", ResponseNamespace:="urn:Microsoft.Search", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function Discovery(registrationXml As String) As String
            Dim array As Object() = Invoke("Discovery", New Object(0) {
                registrationXml
            })
            Return CStr(array(0))
        End Function

        Public Function BeginDiscovery(registrationXml As String, callback As AsyncCallback, asyncState As Object) As IAsyncResult
            Return BeginInvoke("Discovery", New Object(0) {
                registrationXml
            }, callback, asyncState)
        End Function

        Public Function EndDiscovery(asyncResult As IAsyncResult) As String
            Dim array As Object() = EndInvoke(asyncResult)
            Return CStr(array(0))
        End Function

        Public Sub DiscoveryAsync(registrationXml As String)
            DiscoveryAsync(registrationXml, Nothing)
        End Sub

        Public Sub DiscoveryAsync(registrationXml As String, userState As Object)
            If DiscoveryOperationCompleted Is Nothing Then
                DiscoveryOperationCompleted = AddressOf OnDiscoveryOperationCompleted
            End If
            InvokeAsync("Discovery", New Object(0) {
                registrationXml
            }, DiscoveryOperationCompleted, userState)
        End Sub

        Private Sub OnDiscoveryOperationCompleted(arg As Object)
            Dim invokeCompletedEventArgs1 As InvokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
            RaiseEvent DiscoveryCompleted(Me, New DiscoveryCompletedEventArgs(invokeCompletedEventArgs1.Results, invokeCompletedEventArgs1.[Error], invokeCompletedEventArgs1.Cancelled, invokeCompletedEventArgs1.UserState))
        End Sub

        ''' <remarks/>
        Public Shadows Sub CancelAsync(userState As Object)
            MyBase.CancelAsync(userState)
        End Sub
    End Class
End Namespace
