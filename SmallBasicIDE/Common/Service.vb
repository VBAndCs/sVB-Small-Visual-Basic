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
            Set(ByVal value As String)

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
            Set(ByVal value As Boolean)
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

        Public Sub GetCurrentVersionAsync(ByVal userState As Object)
            If GetCurrentVersionOperationCompleted Is Nothing Then
                GetCurrentVersionOperationCompleted = AddressOf OnGetCurrentVersionOperationCompleted
            End If

            InvokeAsync("GetCurrentVersion", New Object(-1) {}, GetCurrentVersionOperationCompleted, userState)
        End Sub

        Private Sub OnGetCurrentVersionOperationCompleted(ByVal arg As Object)
            If GetCurrentVersionCompletedEvent IsNot Nothing Then
                Dim invokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
                RaiseEvent GetCurrentVersionCompleted(Me, New GetCurrentVersionCompletedEventArgs(invokeCompletedEventArgs.Results, invokeCompletedEventArgs.Error, invokeCompletedEventArgs.Cancelled, invokeCompletedEventArgs.UserState))
            End If
        End Sub

        <SoapDocumentMethod("http://tempuri.org/SaveProgram", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function SaveProgram(ByVal title As String, ByVal programText As String, ByVal baseId As String) As String
            Dim array = Invoke("SaveProgram", New Object(2) {title, programText, baseId})
            Return CStr(array(0))
        End Function

        Public Sub SaveProgramAsync(ByVal title As String, ByVal programText As String, ByVal baseId As String)
            SaveProgramAsync(title, programText, baseId, Nothing)
        End Sub

        Public Sub SaveProgramAsync(ByVal title As String, ByVal programText As String, ByVal baseId As String, ByVal userState As Object)
            If SaveProgramOperationCompleted Is Nothing Then
                SaveProgramOperationCompleted = AddressOf OnSaveProgramOperationCompleted
            End If

            InvokeAsync("SaveProgram", New Object(2) {title, programText, baseId}, SaveProgramOperationCompleted, userState)
        End Sub

        Private Sub OnSaveProgramOperationCompleted(ByVal arg As Object)
            If SaveProgramCompletedEvent IsNot Nothing Then
                Dim invokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
                RaiseEvent SaveProgramCompleted(Me, New SaveProgramCompletedEventArgs(invokeCompletedEventArgs.Results, invokeCompletedEventArgs.Error, invokeCompletedEventArgs.Cancelled, invokeCompletedEventArgs.UserState))
            End If
        End Sub

        <SoapDocumentMethod("http://tempuri.org/LoadProgram", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function LoadProgram(ByVal id As String) As String
            Dim array = Invoke("LoadProgram", New Object(0) {id})
            Return CStr(array(0))
        End Function

        Public Sub LoadProgramAsync(ByVal id As String)
            LoadProgramAsync(id, Nothing)
        End Sub

        Public Sub LoadProgramAsync(ByVal id As String, ByVal userState As Object)
            If LoadProgramOperationCompleted Is Nothing Then
                LoadProgramOperationCompleted = AddressOf OnLoadProgramOperationCompleted
            End If

            InvokeAsync("LoadProgram", New Object(0) {id}, LoadProgramOperationCompleted, userState)
        End Sub

        Private Sub OnLoadProgramOperationCompleted(ByVal arg As Object)
            If LoadProgramCompletedEvent IsNot Nothing Then
                Dim invokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
                RaiseEvent LoadProgramCompleted(Me, New LoadProgramCompletedEventArgs(invokeCompletedEventArgs.Results, invokeCompletedEventArgs.Error, invokeCompletedEventArgs.Cancelled, invokeCompletedEventArgs.UserState))
            End If
        End Sub

        <SoapDocumentMethod("http://tempuri.org/PublishProgramDetails", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function PublishProgramDetails(ByVal id As String, ByVal title As String, ByVal description As String, ByVal category As String) As String
            Dim array = Invoke("PublishProgramDetails", New Object(3) {id, title, description, category})
            Return CStr(array(0))
        End Function

        Public Sub PublishProgramDetailsAsync(ByVal id As String, ByVal title As String, ByVal description As String, ByVal category As String)
            PublishProgramDetailsAsync(id, title, description, category, Nothing)
        End Sub

        Public Sub PublishProgramDetailsAsync(ByVal id As String, ByVal title As String, ByVal description As String, ByVal category As String, ByVal userState As Object)
            If PublishProgramDetailsOperationCompleted Is Nothing Then
                PublishProgramDetailsOperationCompleted = AddressOf OnPublishProgramDetailsOperationCompleted
            End If

            InvokeAsync("PublishProgramDetails", New Object(3) {id, title, description, category}, PublishProgramDetailsOperationCompleted, userState)
        End Sub

        Private Sub OnPublishProgramDetailsOperationCompleted(ByVal arg As Object)
            If PublishProgramDetailsCompletedEvent IsNot Nothing Then
                Dim invokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
                RaiseEvent PublishProgramDetailsCompleted(Me, New PublishProgramDetailsCompletedEventArgs(invokeCompletedEventArgs.Results, invokeCompletedEventArgs.Error, invokeCompletedEventArgs.Cancelled, invokeCompletedEventArgs.UserState))
            End If
        End Sub

        <SoapDocumentMethod("http://tempuri.org/GetProgramDetails", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function GetProgramDetails(ByVal id As String) As ProgramDetails
            Dim array = Invoke("GetProgramDetails", New Object(0) {id})
            Return CType(array(0), ProgramDetails)
        End Function

        Public Sub GetProgramDetailsAsync(ByVal id As String)
            GetProgramDetailsAsync(id, Nothing)
        End Sub

        Public Sub GetProgramDetailsAsync(ByVal id As String, ByVal userState As Object)
            If GetProgramDetailsOperationCompleted Is Nothing Then
                GetProgramDetailsOperationCompleted = AddressOf OnGetProgramDetailsOperationCompleted
            End If

            InvokeAsync("GetProgramDetails", New Object(0) {id}, GetProgramDetailsOperationCompleted, userState)
        End Sub

        Private Sub OnGetProgramDetailsOperationCompleted(ByVal arg As Object)
            If GetProgramDetailsCompletedEvent IsNot Nothing Then
                Dim invokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
                RaiseEvent GetProgramDetailsCompleted(Me, New GetProgramDetailsCompletedEventArgs(invokeCompletedEventArgs.Results, invokeCompletedEventArgs.Error, invokeCompletedEventArgs.Cancelled, invokeCompletedEventArgs.UserState))
            End If
        End Sub

        <SoapDocumentMethod("http://tempuri.org/SubmitRating", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=SoapBindingUse.Literal, ParameterStyle:=SoapParameterStyle.Wrapped)>
        Public Function SubmitRating(ByVal id As String, ByVal rating As Double) As ProgramDetails
            Dim array = Invoke("SubmitRating", New Object(1) {id, rating})
            Return CType(array(0), ProgramDetails)
        End Function

        Public Sub SubmitRatingAsync(ByVal id As String, ByVal rating As Double)
            SubmitRatingAsync(id, rating, Nothing)
        End Sub

        Public Sub SubmitRatingAsync(ByVal id As String, ByVal rating As Double, ByVal userState As Object)
            If SubmitRatingOperationCompleted Is Nothing Then
                SubmitRatingOperationCompleted = AddressOf OnSubmitRatingOperationCompleted
            End If

            InvokeAsync("SubmitRating", New Object(1) {id, rating}, SubmitRatingOperationCompleted, userState)
        End Sub

        Private Sub OnSubmitRatingOperationCompleted(ByVal arg As Object)
            If SubmitRatingCompletedEvent IsNot Nothing Then
                Dim invokeCompletedEventArgs = CType(arg, InvokeCompletedEventArgs)
                RaiseEvent SubmitRatingCompleted(Me, New SubmitRatingCompletedEventArgs(invokeCompletedEventArgs.Results, invokeCompletedEventArgs.Error, invokeCompletedEventArgs.Cancelled, invokeCompletedEventArgs.UserState))
            End If
        End Sub

        Public Overloads Sub CancelAsync(ByVal userState As Object)
            MyBase.CancelAsync(userState)
        End Sub

        Private Function IsLocalFileSystemWebService(ByVal url As String) As Boolean
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
