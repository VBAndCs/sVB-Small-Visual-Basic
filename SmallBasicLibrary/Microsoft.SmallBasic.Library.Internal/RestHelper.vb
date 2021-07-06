' To configure or remove Option's included in result, go to Options/Advanced Options...
Option Compare Text
Option Explicit On
Option Infer On
Option Strict On
Imports System.IO
Imports System.Net
Imports System.Threading
Imports System.Xml.Linq

Namespace Microsoft.SmallBasic.Library.Internal
    ''' <summary>
    ''' A private static helper for calling Rest based APIs
    ''' </summary>
    Friend NotInheritable Class RestHelper
        ''' <summary>
        ''' Given a Rest URL, gets the contents as an XML Document.
        ''' </summary>
        Friend Shared Function GetContents(url As String) As XDocument
            Dim result As XDocument = Nothing
            Try
                Dim webRequest1 As WebRequest = WebRequest.Create(url)
                Dim evt As New ManualResetEvent(initialState:=False)
                Dim asyncResult As IAsyncResult = webRequest1.BeginGetResponse(Sub()
                                                                                   evt.[Set]()
                                                                               End Sub, Nothing)
                evt.WaitOne()
                Using webResponse1 As WebResponse = webRequest1.EndGetResponse(asyncResult)
                    Using stream1 As Stream = webResponse1.GetResponseStream()
                        Dim text As String = New StreamReader(stream1).ReadToEnd()
                        result = XDocument.Parse(text)
                        Return result
                    End Using
                End Using
            Catch __unusedException1__ As Exception
                Return result
            End Try
        End Function
    End Class
End Namespace
