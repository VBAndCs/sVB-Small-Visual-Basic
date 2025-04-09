Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Json

Namespace Microsoft.SmallVisualBasic
    <DataContract> Public Class ProgramInfo
        <DataMember> Public Property Title As String = "My App"
        <DataMember> Public Property Description As String = "This app is created by Small Visual Basic."
        <DataMember> Public Property Company As String = "Modern VB"
        <DataMember> Public Property Product As String = "My App"
        <DataMember> Public Property Copyright As String = "Copyright @" & Now.Year
        <DataMember> Public Property Trademark As String = ""
        <DataMember> Public Property Version As String = "1.0.0.0"

        Public Shared Function GetProperties(filePath As String, appName As String) As ProgramInfo
            If File.Exists(filePath) Then
                Try
                    Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                        Dim serializer As New DataContractJsonSerializer(GetType(ProgramInfo))
                        Return CType(serializer.ReadObject(fs), ProgramInfo)
                    End Using
                Catch ex As Exception
                End Try

            Else
                Dim info As New ProgramInfo() With {.Product = appName, .Title = appName}
                ' Writting the Json file manually to apply formatting
                Using writer As New System.IO.StreamWriter(filePath, False)
                    writer.WriteLine("{")
                    writer.WriteLine("    ""Title"": """ & info.Title & """,")
                    writer.WriteLine("    ""Description"": """ & info.Description & """,")
                    writer.WriteLine("    ""Company"": """ & info.Company & """,")
                    writer.WriteLine("    ""Product"": """ & info.Product & """,")
                    writer.WriteLine("    ""Copyright"": """ & info.Copyright & """,")
                    writer.WriteLine("    ""Trademark"": """ & info.Trademark & """,")
                    writer.WriteLine("    ""Version"": """ & info.Version & """")
                    writer.WriteLine("}")
                End Using
                Return info
            End If

            Return New ProgramInfo() With {.Product = appName, .Title = appName}
        End Function
    End Class

End Namespace