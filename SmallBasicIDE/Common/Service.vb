Imports Microsoft.SmallBasic.Properties
Imports System
Imports System.CodeDom.Compiler
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Threading
Imports System.Web.Services
Imports System.Web.Services.Description
Imports System.Web.Services.Protocols

Namespace Microsoft.SmallBasic.com.smallbasic
    <GeneratedCode("System.Web.Services", "4.0.30319.1")>
    <WebServiceBinding(Name:="ServiceSoap", [Namespace]:="http://tempuri.org/")>
    <DebuggerStepThrough>
    <DesignerCategory("code")>
    Public Class Service
        Inherits SoapHttpClientProtocol

        Private GetCurrentVersionOperationCompleted As SendOrPostCallback
        Private SaveProgramOperationCompleted As SendOrPostCallback
        Private LoadProgramOperationCompleted As SendOrPostCallback
        Private PublishProgramDetailsOperationCompleted As SendOrPostCallback
        Private GetProgramDetailsOperationCompleted As SendOrPostCallback
        Private SubmitRatingOperationCompleted As SendOrPostCallback
        Private useDefaultCredentialsSetExplicitly As Boolean

        Public Overloads Property Url As String
            Get
                Return MyBase.Url
            End Get
            Set(value As String)

                If IsLocalFileSystemWebService(MyBase.Url) AndAlso Not useDefaultCredentialsSetExplicitly AndAlso Not IsLocalFileSystemWebService(value) Then
                    MyBase.UseDefaultCredentials = False
                End If

                MyBase.Url = value
            End Set
        End Property

        Public Overloads Property UseDefaultCredentials As Boolean
            Get
                Return MyBase.UseDefaultCredentials
            End Get
            Set(value As Boolean)
                MyBase.UseDefaultCredentials = value
                useDefaultCredentialsSetExplicitly = True
            End Set
        End Property

        Public Event GetCurrentVersionCompleted As GetCurrentVersionCompletedEventHandler
        Public Event SaveProgramCompleted As SaveProgramCompletedEventHandler
        Public Event LoadProgramCompleted As LoadProgramCompletedEventHandler
        Public Event PublishProgramDetailsCompleted As PublishProgramDetailsCompletedEventHandler
        Public Event GetProgramDetailsCompleted As GetProgramDetailsCompletedEventHandler
        Public Event SubmitRatingCompleted As SubmitRatingCompletedEventHandler

        Public Sub New()
            Url = Settings.Default.SB_com_smallbasic_Service

            If IsLocalFileSystemWebService(Url) Then
                UseDefaultCredentials = True
                useDefaultCredentialsSetExplicitly = False
            Else
                useDefaultCredentialsSetExplicitly = True
            End If
        End Sub

        <SoapDocumentMethod("http://tempuri.org/GetCurrentVersion", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function GetCurrentVersion() As String
            Dim array = Invoke("GetCurrentVersion", New Object(-1) {})
            Return CStr(array(0))
        End Function

        Public Sub GetCurrentVersionAsync()
            GetCurrentVersionAsync(Nothing)
        End Sub

        Public Sub GetCurrentVersionAsync(userState As Object)
            If GetCurrentVersionOperationCompleted Is Nothing Then
                GetCurrentVersionOperationCompleted = AddressOf OnGetCurrentVersionOperationCompleted
            End If

            InvokeAsync("GetCurrentVersion", New Object(-1) {}, GetCurrentVersionOperationCompleted, userState)
        End Sub

        Private Sub OnGetCurrentVersionOperationCompleted(arg As Object)
            If GetCurrentVersionCompletedEvent IsNot Nothing Then
                Dim invokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
                RaiseEvent GetCurrentVersionCompleted(Me, New GetCurrentVersionCompletedEventArgs(invokeCompletedEventArgs.Results, invokeCompletedEventArgs.Error, invokeCompletedEventArgs.Cancelled, invokeCompletedEventArgs.UserState))
            End If
        End Sub

        <SoapDocumentMethod("http://tempuri.org/SaveProgram", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function SaveProgram(title As String, programText As String, baseId As String) As String
            Dim array = Invoke("SaveProgram", New Object(2) {title, programText, baseId})
            Return CStr(array(0))
        End Function

        Public Sub SaveProgramAsync(title As String, programText As String, baseId As String)
            SaveProgramAsync(title, programText, baseId, Nothing)
        End Sub

        Public Sub SaveProgramAsync(title As String, programText As String, baseId As String, userState As Object)
            If SaveProgramOperationCompleted Is Nothing Then
                SaveProgramOperationCompleted = AddressOf OnSaveProgramOperationCompleted
            End If

            InvokeAsync("SaveProgram", New Object(2) {title, programText, baseId}, SaveProgramOperationCompleted, userState)
        End Sub

        Private Sub OnSaveProgramOperationCompleted(arg As Object)
            If SaveProgramCompletedEvent IsNot Nothing Then
                Dim invokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
                RaiseEvent SaveProgramCompleted(Me, New SaveProgramCompletedEventArgs(invokeCompletedEventArgs.Results, invokeCompletedEventArgs.Error, invokeCompletedEventArgs.Cancelled, invokeCompletedEventArgs.UserState))
            End If
        End Sub

        <SoapDocumentMethod("http://tempuri.org/LoadProgram", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function LoadProgram(id As String) As String
            Dim array = Invoke("LoadProgram", New Object(0) {id})
            Return CStr(array(0))
        End Function

        Public Sub LoadProgramAsync(id As String)
            LoadProgramAsync(id, Nothing)
        End Sub

        Public Sub LoadProgramAsync(id As String, userState As Object)
            If LoadProgramOperationCompleted Is Nothing Then
                LoadProgramOperationCompleted = AddressOf OnLoadProgramOperationCompleted
            End If

            InvokeAsync("LoadProgram", New Object(0) {id}, LoadProgramOperationCompleted, userState)
        End Sub

        Private Sub OnLoadProgramOperationCompleted(arg As Object)
            If LoadProgramCompletedEvent IsNot Nothing Then
                Dim invokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
                RaiseEvent LoadProgramCompleted(Me, New LoadProgramCompletedEventArgs(invokeCompletedEventArgs.Results, invokeCompletedEventArgs.Error, invokeCompletedEventArgs.Cancelled, invokeCompletedEventArgs.UserState))
            End If
        End Sub

        <SoapDocumentMethod("http://tempuri.org/PublishProgramDetails", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function PublishProgramDetails(id As String, title As String, description As String, category As String) As String
            Dim array = Invoke("PublishProgramDetails", New Object(3) {id, title, description, category})
            Return CStr(array(0))
        End Function

        Public Sub PublishProgramDetailsAsync(id As String, title As String, description As String, category As String)
            PublishProgramDetailsAsync(id, title, description, category, Nothing)
        End Sub

        Public Sub PublishProgramDetailsAsync(id As String, title As String, description As String, category As String, userState As Object)
            If PublishProgramDetailsOperationCompleted Is Nothing Then
                PublishProgramDetailsOperationCompleted = AddressOf OnPublishProgramDetailsOperationCompleted
            End If

            InvokeAsync("PublishProgramDetails", New Object(3) {id, title, description, category}, PublishProgramDetailsOperationCompleted, userState)
        End Sub

        Private Sub OnPublishProgramDetailsOperationCompleted(arg As Object)
            If PublishProgramDetailsCompletedEvent IsNot Nothing Then
                Dim invokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
                RaiseEvent PublishProgramDetailsCompleted(Me, New PublishProgramDetailsCompletedEventArgs(invokeCompletedEventArgs.Results, invokeCompletedEventArgs.Error, invokeCompletedEventArgs.Cancelled, invokeCompletedEventArgs.UserState))
            End If
        End Sub

        <SoapDocumentMethod("http://tempuri.org/GetProgramDetails", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function GetProgramDetails(id As String) As ProgramDetails
            Dim array = Invoke("GetProgramDetails", New Object(0) {id})
            Return CType(array(0), ProgramDetails)
        End Function

        Public Sub GetProgramDetailsAsync(id As String)
            GetProgramDetailsAsync(id, Nothing)
        End Sub

        Public Sub GetProgramDetailsAsync(id As String, userState As Object)
            If GetProgramDetailsOperationCompleted Is Nothing Then
                GetProgramDetailsOperationCompleted = AddressOf OnGetProgramDetailsOperationCompleted
            End If

            InvokeAsync("GetProgramDetails", New Object(0) {id}, GetProgramDetailsOperationCompleted, userState)
        End Sub

        Private Sub OnGetProgramDetailsOperationCompleted(arg As Object)
            If GetProgramDetailsCompletedEvent IsNot Nothing Then
                Dim invokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
                RaiseEvent GetProgramDetailsCompleted(Me, New GetProgramDetailsCompletedEventArgs(invokeCompletedEventArgs.Results, invokeCompletedEventArgs.Error, invokeCompletedEventArgs.Cancelled, invokeCompletedEventArgs.UserState))
            End If
        End Sub

        <SoapDocumentMethod("http://tempuri.org/SubmitRating", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function SubmitRating(id As String, rating As Double) As ProgramDetails
            Dim array = Invoke("SubmitRating", New Object(1) {id, rating})
            Return CType(array(0), ProgramDetails)
        End Function

        Public Sub SubmitRatingAsync(id As String, rating As Double)
            SubmitRatingAsync(id, rating, Nothing)
        End Sub

        Public Sub SubmitRatingAsync(id As String, rating As Double, userState As Object)
            If SubmitRatingOperationCompleted Is Nothing Then
                SubmitRatingOperationCompleted = AddressOf OnSubmitRatingOperationCompleted
            End If

            InvokeAsync("SubmitRating", New Object(1) {id, rating}, SubmitRatingOperationCompleted, userState)
        End Sub

        Private Sub OnSubmitRatingOperationCompleted(arg As Object)
            If SubmitRatingCompletedEvent IsNot Nothing Then
                Dim invokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
                RaiseEvent SubmitRatingCompleted(Me, New SubmitRatingCompletedEventArgs(invokeCompletedEventArgs.Results, invokeCompletedEventArgs.Error, invokeCompletedEventArgs.Cancelled, invokeCompletedEventArgs.UserState))
            End If
        End Sub

        Public Overloads Sub CancelAsync(userState As Object)
            MyBase.CancelAsync(userState)
        End Sub

        Private Function IsLocalFileSystemWebService(url As String) As Boolean
            If Equals(url, Nothing) OrElse Equals(url, String.Empty) Then
                Return False
            End If

            Dim uri As Uri = New Uri(url)

            If uri.Port >= 1024 AndAlso String.Compare(uri.Host, "localHost", StringComparison.OrdinalIgnoreCase) = 0 Then
                Return True
            End If

            Return False
        End Function
    End Class
End Namespace
