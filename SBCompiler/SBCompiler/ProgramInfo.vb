Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Json

<DataContract>
Public Class ProgramInfo
    <DataMember> Public Property Title As String = "My App"
    <DataMember> Public Property Description As String = "This app is created by Small Visual Basic."
    <DataMember> Public Property Company As String = "Modern VB"
    <DataMember> Public Property Product As String = "My App"
    <DataMember> Public Property Copyright As String = "Copyright @" & Now.Year
    <DataMember> Public Property Trademark As String
    <DataMember> Public Property Version As String = "1.0.0.0"

    Public Shared Function GetProperties(filePath As String) As ProgramInfo
        If File.Exists(filePath) Then
            Try
                Dim serializer As New DataContractJsonSerializer(GetType(ProgramInfo))
                Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                    Return CType(serializer.ReadObject(fs), ProgramInfo)
                End Using
            Catch ex As Exception
            End Try
        End If

        Return New ProgramInfo
    End Function
End Class

