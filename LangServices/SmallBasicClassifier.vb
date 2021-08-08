Imports System.Windows.Threading
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Classification
Imports sb = Microsoft.SmallBasic

Namespace Microsoft.SmallBasic.LanguageService
    Public NotInheritable Class SmallBasicClassifier
        Implements IClassifier

        Private textBuffer As ITextBuffer
        Private classificationTypeRegistry As IClassificationTypeRegistry
        Private keywordType As IClassificationType
        Private textType As IClassificationType
        Private commentType As IClassificationType
        Private operatorType As IClassificationType
        Private literalType As IClassificationType
        Private numericType As IClassificationType
        Private stringType As IClassificationType
        Private identifierType As IClassificationType
        Private typeType As IClassificationType
        Private memberType As IClassificationType
        Private delimiterType As IClassificationType
        Private unknownType As IClassificationType
        Private compiler As Compiler
        Public Event ClassificationChanged As EventHandler(Of ClassificationChangedEventArgs) Implements IClassifier.ClassificationChanged

        Public Sub New(ByVal textBuffer As ITextBuffer, ByVal classificationTypeRegistry As IClassificationTypeRegistry)
            If Not textBuffer.Properties.TryGetProperty(GetType(Compiler), compiler) Then
                compiler = New Compiler()
                textBuffer.Properties.AddProperty(GetType(Compiler), compiler)
            End If

            Me.textBuffer = textBuffer
            Me.classificationTypeRegistry = classificationTypeRegistry
            keywordType = classificationTypeRegistry.GetClassificationType("SB_Keyword")
            textType = classificationTypeRegistry.GetClassificationType("SB_Text")
            commentType = classificationTypeRegistry.GetClassificationType("SB_Comment")
            operatorType = classificationTypeRegistry.GetClassificationType("SB_Operator")
            literalType = classificationTypeRegistry.GetClassificationType("SB_Literal")
            numericType = classificationTypeRegistry.GetClassificationType("SB_Numeric")
            stringType = classificationTypeRegistry.GetClassificationType("SB_String")
            identifierType = classificationTypeRegistry.GetClassificationType("SB_Identifier")
            typeType = classificationTypeRegistry.GetClassificationType("SB_Type")
            memberType = classificationTypeRegistry.GetClassificationType("SB_Member")
            delimiterType = classificationTypeRegistry.GetClassificationType("SB_Delimiter")
            unknownType = classificationTypeRegistry.GetClassificationType("SB_Unknown")
        End Sub

        Private Sub OnUnitUpdated(ByVal sender As Object, ByVal e As EventArgs)
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, CType(Function()
                                                                                       If ClassificationChangedEvent IsNot Nothing Then
                                                                                           Dim changeSpan = textBuffer.CurrentSnapshot.CreateTextSpan(0, textBuffer.CurrentSnapshot.Length, SpanTrackingMode.EdgeInclusive)
                                                                                           RaiseEvent ClassificationChanged(Me, New ClassificationChangedEventArgs(changeSpan))
                                                                                       End If

                                                                                       Return Nothing
                                                                                   End Function, DispatcherOperationCallback), Nothing)
        End Sub

        Private Function GetClassificationForType(ByVal tokenInfo As TokenInfo, ByVal previousToken As TokenInfo, ByVal previousPreviousToken As TokenInfo) As IClassificationType
            Dim value As TypeInfo = Nothing

            If tokenInfo.TokenType = TokenType.Identifier Then
                Dim normalizedText = tokenInfo.NormalizedText

                If compiler.TypeInfoBag.Types.ContainsKey(normalizedText) Then
                    Return typeType
                End If

                If previousToken.Token = Token.Dot AndAlso previousPreviousToken.TokenType = TokenType.Identifier AndAlso compiler.TypeInfoBag.Types.TryGetValue(previousPreviousToken.NormalizedText, value) AndAlso (value.Events.ContainsKey(normalizedText) OrElse value.Methods.ContainsKey(normalizedText) OrElse value.Properties.ContainsKey(normalizedText)) Then
                    Return memberType
                End If

                Return identifierType
            End If

            Select Case tokenInfo.TokenType
                Case TokenType.Comment
                    Return commentType
                Case TokenType.Delimiter
                    Return delimiterType
                Case TokenType.Identifier
                    Return identifierType
                Case TokenType.Illegal
                    Return textType
                Case TokenType.Keyword
                    Return keywordType
                Case TokenType.NumericLiteral
                    Return numericType
                Case TokenType.Operator
                    Return operatorType
                Case TokenType.StringLiteral
                    Return stringType
                Case Else
                    Return textType
            End Select
        End Function

        Public Function GetClassificationSpans(ByVal textSpan As SnapshotSpan) As IList(Of ClassificationSpan) Implements IClassifier.GetClassificationSpans
            Dim list As List(Of ClassificationSpan) = New List(Of ClassificationSpan)()
            Dim lineNumberFromPosition = textSpan.Snapshot.GetLineNumberFromPosition(textSpan.Start)
            Dim lineNumberFromPosition2 = textSpan.Snapshot.GetLineNumberFromPosition(textSpan.End)

            For i = lineNumberFromPosition To lineNumberFromPosition2
                Dim lineFromLineNumber = textSpan.Snapshot.GetLineFromLineNumber(i)
                Dim tokenList As TokenEnumerator = lineScanner.GetTokenEnumerator(lineFromLineNumber.GetText(), i)
                Dim tokenInfo = sb.TokenInfo.Illegal
                Dim tokenInfo2 = sb.TokenInfo.Illegal
                Dim illegal = TokenInfo.Illegal

                Do
                    illegal = tokenInfo2
                    tokenInfo2 = tokenInfo
                    tokenInfo = tokenList.Current
                    Dim classificationForType = GetClassificationForType(tokenInfo, tokenInfo2, illegal)
                    Dim textSpan2 = textSpan.Snapshot.CreateTextSpan(lineFromLineNumber.Start + tokenInfo.Column, If(Not Equals(tokenInfo.Text, Nothing), tokenInfo.Text.Length, 0), SpanTrackingMode.EdgeInclusive)
                    list.Add(New ClassificationSpan(textSpan2, classificationForType))
                Loop While tokenList.MoveNext()
            Next

            Return list
        End Function
    End Class
End Namespace
