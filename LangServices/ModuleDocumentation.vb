Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Xml

Namespace Microsoft.SmallBasic.LanguageService
    Public Class ModuleDocumentation
        Private _itemDocMap As Dictionary(Of String, CompletionItemDocumentation)

        Public Sub New(modulePath As String)
            Dim localizedModuleDocPath = GetLocalizedModuleDocPath(modulePath)

            If File.Exists(localizedModuleDocPath) Then
                ProcessDocumentation(localizedModuleDocPath)
            End If
        End Sub

        Public Function GetItemDocumentation(itemName As String) As CompletionItemDocumentation
            Dim value As CompletionItemDocumentation = Nothing

            If _itemDocMap IsNot Nothing Then
                _itemDocMap.TryGetValue(itemName, value)
            End If

            Return value
        End Function

        Private Function GetLocalizedModuleDocPath(modulePath As String) As String
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

        Private Sub ProcessDocumentation(xmlFilePath As String)
            _itemDocMap = New Dictionary(Of String, CompletionItemDocumentation)()

            Try
                Dim xmlDocument As XmlDocument = New XmlDocument()
                xmlDocument.PreserveWhitespace = False
                xmlDocument.Load(xmlFilePath)
                Dim xmlNodeList = xmlDocument.SelectNodes("doc/members/member")

                For Each item As XmlNode In xmlNodeList
                    Dim xmlAttribute = item.Attributes("name")

                    If xmlAttribute Is Nothing Then Continue For

                    Dim value = xmlAttribute.Value
                    Dim documentation As New CompletionItemDocumentation()
                    _itemDocMap(value) = documentation

                    Dim xmlNode2 = item.SelectSingleNode("summary")
                    documentation.Summary = GetTextFromXmlNode(xmlNode2)

                    Dim xmlNode3 = item.SelectSingleNode("returns")
                    documentation.Returns = GetTextFromXmlNode(xmlNode3)

                    Dim xmlNode4 = item.SelectSingleNode("example")
                    documentation.Example = GetTextFromXmlNode(xmlNode4)

                    Dim xmlNodeList2 = item.SelectNodes("param")
                    If xmlNodeList2 Is Nothing Then Continue For

                    For Each item2 As XmlNode In xmlNodeList2
                        Dim xmlAttribute2 = item2.Attributes("name")

                        If xmlAttribute2 IsNot Nothing Then
                            Dim value2 = xmlAttribute2.Value
                            documentation.ParamsDoc(value2) = GetTextFromXmlNode(item2)
                        End If
                    Next
                Next

            Catch
            End Try
        End Sub

        Private Function GetTextFromXmlNode(xmlNode As XmlNode) As String
            If xmlNode Is Nothing Then Return Nothing

            Dim stringBuilder As New StringBuilder()
            Dim stringReader As New StringReader(xmlNode.InnerText.Trim())

            Do
                Dim text As String = stringReader.ReadLine()
                If text Is Nothing Then Exit Do

                If text.StartsWith("            ") Then text = text.Substring(12)
                stringBuilder.AppendLine(text)
            Loop

            Return stringBuilder.ToString().TrimEnd
        End Function
    End Class
End Namespace
