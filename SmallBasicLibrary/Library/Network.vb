Imports System.Globalization
Imports System.IO
Imports System.Net

Namespace Library
    ''' <summary>
    ''' This helper class provides network access methods
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class Network
        ''' <summary>
        ''' Downloads a file from the network to a local temporary file.
        ''' </summary>
        ''' <param name="url">
        ''' The URL of the file on the network.
        ''' </param>
        ''' <returns>
        ''' A local file name that the remote file was downloaded as.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function DownloadFile(url As Primitive) As Primitive
            If url.IsEmpty Then
                Return url
            End If
            Dim tempFileName As String = Path.GetTempFileName()
            Dim stream1 As Stream = Nothing
            Dim stream2 As Stream = Nothing
            Dim webResponse1 As WebResponse = Nothing
            Try
                Dim webRequest1 As WebRequest = WebRequest.Create(url)
                webResponse1 = webRequest1.GetResponse()
                stream1 = IO.File.Open(tempFileName, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read)
                Dim buffer As Byte() = New Byte(16383) {}
                Dim num As Long = webResponse1.ContentLength
                stream2 = webResponse1.GetResponseStream()
                While num > 0
                    Dim num2 As Integer = stream2.Read(buffer, 0, 16384)
                    stream1.Write(buffer, 0, num2)
                    num -= num2
                End While
            Catch __unusedException1__ As Exception
                Return ""
            Finally
                stream1?.Close()
                stream2?.Close()
                webResponse1?.Close()
            End Try

            Return tempFileName
        End Function

        ''' <summary>
        ''' Gets the contents of a specified web page.
        ''' </summary>
        ''' <param name="url">
        ''' The URL of the web page
        ''' </param>
        ''' <returns>
        ''' The contents of the specified web page.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetWebPageContents(url As Primitive) As Primitive
            If url.IsEmpty Then
                Return url
            End If

            Return GetWebPageContents(CStr(url))
        End Function

        Friend Shared Function GetLocalFile(fileNameOrUrl As Primitive) As Primitive
            If fileNameOrUrl.IsEmpty Then
                Return fileNameOrUrl
            End If

            Dim result As Uri = Nothing
            If Uri.TryCreate(fileNameOrUrl, UriKind.RelativeOrAbsolute, result) AndAlso result.IsAbsoluteUri Then
                If result.Scheme.ToLower(CultureInfo.InvariantCulture) = "file" Then
                    Return fileNameOrUrl
                End If

                Return DownloadFile(fileNameOrUrl)
            End If

            Return fileNameOrUrl
        End Function

        Friend Shared Function GetWebPageContents(url As String) As String
            Dim streamReader1 As StreamReader = Nothing
            Dim webResponse1 As WebResponse = Nothing
            Dim result As String = ""
            Try
                Dim webRequest1 As WebRequest = WebRequest.Create(url)
                webResponse1 = webRequest1.GetResponse()
                streamReader1 = New StreamReader(webResponse1.GetResponseStream())
                result = streamReader1.ReadToEnd()
                Return result
            Catch __unusedException1__ As Exception
                Return result
            Finally
                streamReader1?.Close()
                webResponse1?.Close()
            End Try
        End Function
    End Class
End Namespace
