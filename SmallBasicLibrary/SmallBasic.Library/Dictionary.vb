' To configure or remove Option's included in result, go to Options/Advanced Options...
Option Compare Text
Option Explicit On
Option Infer On
Option Strict On
Imports System.IO
Imports System.Text
Imports System.Xml
Imports Microsoft.SmallBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' Provides access to an online Dictionary service.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Dictionary
        Private Const _queryXml As String = "<QueryPacket xmlns='urn:Microsoft.Search.Query' revision='1' >" & vbCrLf & "                                   <Query domain='{2}'>" & vbCrLf & "                                     <Context>" & vbCrLf & "                                       <QueryText type='STRING' language='en-us' >{0}</QueryText>" & vbCrLf & "                                       <LanguagePreference>en-us</LanguagePreference>" & vbCrLf & "                                     </Context>" & vbCrLf & "                                     <OfficeContext xmlns='urn:Microsoft.Search.Query.Office.Context' revision='1'>" & vbCrLf & "                                       <UserPreferences>" & vbCrLf & "                                         <ParentalControl>false</ParentalControl>" & vbCrLf & "                                       </UserPreferences>" & vbCrLf & "                                       <ServiceData>{1}</ServiceData>" & vbCrLf & "                                       <ApplicationContext>" & vbCrLf & "                                         <Name>Microsoft Office Word</Name>" & vbCrLf & "                                         <Version>(14.0.3524)</Version>" & vbCrLf & "                                       </ApplicationContext>" & vbCrLf & "                                       <QueryLanguage>en-us</QueryLanguage>" & vbCrLf & "                                       <KeyboardLanguage>en-us</KeyboardLanguage>" & vbCrLf & "                                    </OfficeContext>" & vbCrLf & "                                   </Query>" & vbCrLf & "                                 </QueryPacket>"

        ''' <summary>
        ''' Gets the definition of a word in English.  The same as GetDefinitionEnglishToEnglish.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinition(word As Primitive) As Primitive
            Return GetDefinitionEnglishToEnglish(word)
        End Function

        ''' <summary>
        ''' Gets the definition of a word, English to English.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionEnglishToEnglish(word As Primitive) As Primitive
            Return GetDefinition(word, "EDICT", "{FEF89077-4F4D-4803-A8BF-228083F70EAA}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, Simplified Chinese to English.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionSimplifiedChineseToEnglish(word As Primitive) As Primitive
            Return GetDefinition(word, "2052/1033/4", "{03D4CECF-0578-4DBF-BDEB-F55502662B51}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, English to German.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionEnglishToGerman(word As Primitive) As Primitive
            Return GetDefinition(word, "1033/1031/4", "0FDEA0C3-A12B-4BD8-983E-96BBAE93B097")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, English to Simplified Chinese.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionEnglishToSimplifiedChinese(word As Primitive) As Primitive
            Return GetDefinition(word, "1033/2052/4", "{2D4205BC-1D59-43B4-9156-19F8B8AEF1A4}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, English to French.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionEnglishToFrench(word As Primitive) As Primitive
            Return GetDefinition(word, "1033/1036/4", "{30C1B7BD-93FF-44A6-928F-D848AEE0BDD6}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, German to English.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionGermanToEnglish(word As Primitive) As Primitive
            Return GetDefinition(word, "1031/1033/4", "{63D351DB-D12E-448B-8541-9F794E1EC973}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, English to Japanese.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionEnglishToJapanese(word As Primitive) As Primitive
            Return GetDefinition(word, "1033/1041/4", "{871E52A8-6A33-4753-B3D3-A1A6A723BC9F}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, French to English.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionFrenchToEnglish(word As Primitive) As Primitive
            Return GetDefinition(word, "1036/1033/4", "{FE5F4005-127F-4885-8366-68CF00CED317}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, Itlian to English.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionItalianToEnglish(word As Primitive) As Primitive
            Return GetDefinition(word, "1040/1033/4", "{A7F03864-B9E3-4824-9E30-D0BD52AE37DF}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, Traditional Chinese to English.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionTraditionalChineseToEnglish(word As Primitive) As Primitive
            Return GetDefinition(word, "1028/1033/4", "{9FC92720-10A4-47FE-8066-E7567FBA0AB4}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, Spanish to English.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionSpanishToEnglish(word As Primitive) As Primitive
            Return GetDefinition(word, "3082/1033/4", "{B7B7DDE2-AFFB-415B-BB77-BC1A0D4D694D}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, Japanese to English.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionJapaneseToEnglish(word As Primitive) As Primitive
            Return GetDefinition(word, "1041/1033/4", "{C6032F4A-557E-4FC0-AFA8-46B918B626EB}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, English to Korean.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionEnglishToKorean(word As Primitive) As Primitive
            Return GetDefinition(word, "1033/1042/4", "{CF56F86E-4DBE-42BF-BB08-1C9BBF2B26F2}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, English to Traditional Chinese.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionEnglishToTraditionalChinese(word As Primitive) As Primitive
            Return GetDefinition(word, "1033/1028/4", "{F08B2224-AAA5-42E9-B668-A448BFC16D5B}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, Korean to English.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionKoreanToEnglish(word As Primitive) As Primitive
            Return GetDefinition(word, "1042/1033/4", "{F1A3E34D-7A0A-4118-84F2-17D2CAB58F18}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, English to Italian.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionEnglishToItalian(word As Primitive) As Primitive
            Return GetDefinition(word, "1033/1040/4", "{F8EB68DA-29AF-4CE1-B188-3D5384A7259A}")
        End Function

        ''' <summary>
        ''' Gets the definition of a word, English to Spanish.
        ''' </summary>
        ''' <param name="word">
        ''' The word to define.
        ''' </param>
        ''' <returns>
        ''' The definition(s) of the specified word.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetDefinitionEnglishToSpanish(word As Primitive) As Primitive
            Return GetDefinition(word, "1033/3082/4", "{FE0FEFA5-ED3B-4D8E-AC1A-A308A9210C34}")
        End Function

        Private Shared Function GetDefinition(word As String, serviceCode As String, langGuid As String) As String
            Dim stringBuilder1 As New StringBuilder
            Try
                Dim officeResearch1 As New OfficeResearch
                Dim s As String = officeResearch1.Query(String.Format("<QueryPacket xmlns='urn:Microsoft.Search.Query' revision='1' >" & vbCrLf & "                                   <Query domain='{2}'>" & vbCrLf & "                                     <Context>" & vbCrLf & "                                       <QueryText type='STRING' language='en-us' >{0}</QueryText>" & vbCrLf & "                                       <LanguagePreference>en-us</LanguagePreference>" & vbCrLf & "                                     </Context>" & vbCrLf & "                                     <OfficeContext xmlns='urn:Microsoft.Search.Query.Office.Context' revision='1'>" & vbCrLf & "                                       <UserPreferences>" & vbCrLf & "                                         <ParentalControl>false</ParentalControl>" & vbCrLf & "                                       </UserPreferences>" & vbCrLf & "                                       <ServiceData>{1}</ServiceData>" & vbCrLf & "                                       <ApplicationContext>" & vbCrLf & "                                         <Name>Microsoft Office Word</Name>" & vbCrLf & "                                         <Version>(14.0.3524)</Version>" & vbCrLf & "                                       </ApplicationContext>" & vbCrLf & "                                       <QueryLanguage>en-us</QueryLanguage>" & vbCrLf & "                                       <KeyboardLanguage>en-us</KeyboardLanguage>" & vbCrLf & "                                    </OfficeContext>" & vbCrLf & "                                   </Query>" & vbCrLf & "                                 </QueryPacket>", word, serviceCode, langGuid))
                Dim xmlTextReader1 As New XmlTextReader(New StringReader(s))
                xmlTextReader1.WhitespaceHandling = WhitespaceHandling.Significant
                If xmlTextReader1.ReadToDescendant("Content") Then
                    Dim xmlReader1 As XmlReader = xmlTextReader1.ReadSubtree()
                    While xmlReader1.Read()
                        Select Case xmlReader1.Name
                            Case "Heading", "Line", "P"
                                stringBuilder1.AppendLine()
                            Case "Char", "Text"
                                If xmlReader1.NodeType = XmlNodeType.Element Then
                                    stringBuilder1.Append(" ")
                                End If

                            Case Else

                                If xmlReader1.NodeType = XmlNodeType.Text Then
                                    stringBuilder1.Append(xmlReader1.Value)
                                End If
                        End Select
                    End While
                End If
            Catch __unusedException1__ As Exception

            End Try

            Return stringBuilder1.ToString()
        End Function
    End Class
End Namespace
