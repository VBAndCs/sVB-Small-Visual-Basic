Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Xml

Namespace Microsoft.SmallBasic.LanguageService
    Public Class ModuleDocumentation
        Private _itemDocMap As Dictionary(Of String, CompletionItemDocumentation)

        Public Sub New(ByVal modulePath As String)
            Dim localizedModuleDocPath = GetLocalizedModuleDocPath(modulePath)

            If File.Exists(localizedModuleDocPath) Then
                ProcessDocumentation(localizedModuleDocPath)
            End If
        End Sub

        Public Function GetItemDocumentation(ByVal itemName As String) As CompletionItemDocumentation
            Dim value As CompletionItemDocumentation = Nothing

            If _itemDocMap IsNot Nothing Then
                _itemDocMap.TryGetValue(itemName, value)
            End If

            Return value
        End Function

        Private Function GetLocalizedModuleDocPath(ByVal modulePath As String) As String
            Dim directoryName = Path.GetDirectoryName(modulePath)
            Dim fileNameWithoutExtension = Path.GetFileNameWithoutExtension(modulePath)
            Dim ietfLanguageTag = CultureInfo.CurrentUICulture.IetfLanguageTag
            Dim text = Path.Combine(directoryName, $"{fileNameWithoutExtension}.{ietfLanguageTag}.xml")

            If File.Exists(text) Then
                Return text
            End If

            ietfLanguageTag = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName
            text = Path.Combine(directoryName, $"{fileNameWithoutExtension}.{ietfLanguageTag}.xml")

            If File.Exists(text) Then
                Return text
            End If

            If CultureInfo.CurrentUICulture.Parent IsNot Nothing Then
                ietfLanguageTag = CultureInfo.CurrentUICulture.Parent.IetfLanguageTag
                text = Path.Combine(directoryName, $"{fileNameWithoutExtension}.{ietfLanguageTag}.xml")

                If File.Exists(text) Then
                    Return text
                End If
            End If

            Return Path.Combine(directoryName, fileNameWithoutExtension & ".xml")
        End Function

        Private Sub ProcessDocumentation(ByVal xmlFilePath As String)
            _itemDocMap = New Dictionary(Of String, CompletionItemDocumentation)()

            Try
                Dim xmlDocument As XmlDocument = New XmlDocument()
                xmlDocument.PreserveWhitespace = False
                xmlDocument.Load(xmlFilePath)
                Dim xmlNodeList = xmlDocument.SelectNodes("doc/members/member")

                For Each item As XmlNode In xmlNodeList
                    Dim xmlAttribute = item.Attributes("name")

                    If xmlAttribute Is Nothing Then
                        Continue For
                    End If

                    Dim value = xmlAttribute.Value
                    Dim completionItemDocumentation As CompletionItemDocumentation = New CompletionItemDocumentation()
                    _itemDocMap(value) = completionItemDocumentation
                    Dim xmlNode2 = item.SelectSingleNode("summary")
                    completionItemDocumentation.Summary = GetTextFromXmlNode(xmlNode2)
                    Dim xmlNode3 = item.SelectSingleNode("returns")
                    completionItemDocumentation.Returns = GetTextFromXmlNode(xmlNode3)
                    Dim xmlNode4 = item.SelectSingleNode("example")
                    completionItemDocumentation.Example = GetTextFromXmlNode(xmlNode4)
                    Dim xmlNodeList2 = item.SelectNodes("param")

                    If xmlNodeList2 Is Nothing Then
                        Continue For
                    End If

                    For Each item2 As XmlNode In xmlNodeList2
                        Dim xmlAttribute2 = item2.Attributes("name")

                        If xmlAttribute2 IsNot Nothing Then
                            Dim value2 = xmlAttribute2.Value
                            completionItemDocumentation.ParamsDoc(value2) = GetTextFromXmlNode(item2)
                        End If
                    Next
                Next

            Catch
            End Try
        End Sub

        Private Function GetTextFromXmlNode(ByVal xmlNode As XmlNode) As String
            If xmlNode Is Nothing Then
                Return Nothing
            End If

            Dim stringBuilder As StringBuilder = New StringBuilder()
            Dim stringReader As StringReader = New StringReader(xmlNode.InnerText.Trim())

            While True
                Dim text As String = stringReader.ReadLine()

                If Equals(text, Nothing) Then
                    Exit While
                End If

                If text.StartsWith("            ") Then
                    text = text.Substring(12)
                End If

                stringBuilder.AppendLine(text)
            End While

            Return stringBuilder.ToString()
        End Function
    End Class
End Namespace
