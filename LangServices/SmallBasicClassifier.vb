Imports System.Windows.Threading
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Classification
Imports sb = Microsoft.SmallVisualBasic

Namespace Microsoft.SmallVisualBasic.LanguageService
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

        Public Sub New(
                   textBuffer As ITextBuffer,
                   classificationTypeRegistry As IClassificationTypeRegistry
                )

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

        Private Sub OnUnitUpdated(sender As Object, e As EventArgs)
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal,
                CType(Function()
                          If ClassificationChangedEvent IsNot Nothing Then
                              Dim snapshot = textBuffer.CurrentSnapshot
                              Dim changeSpan = snapshot.CreateTextSpan(
                                      0, snapshot.Length,
                                      SpanTrackingMode.EdgeInclusive
                              )
                              RaiseEvent ClassificationChanged(Me,
                                       New ClassificationChangedEventArgs(changeSpan))
                          End If

                          Return Nothing
                      End Function, DispatcherOperationCallback), Nothing)
        End Sub

        Private Function GetClassificationForType(
                        token As Token,
                        prevToken As Token,
                        b4PrevToken As Token
                    ) As IClassificationType


            Select Case token.Type
                Case TokenType.Dot, TokenType.Lookup
                    Return keywordType
            End Select

            If token.ParseType = ParseType.Identifier Then
                Dim text = token.LCaseText

                If text = "me" OrElse text = "global" Then
                    Return keywordType
                End If

                If (prevToken.Type = TokenType.Dot OrElse prevToken.Type = TokenType.Lookup) Then
                    If b4PrevToken.ParseType = ParseType.Identifier Then
                        Return memberType
                    End If
                End If

                If compiler.TypeInfoBag.Types.ContainsKey(text) Then
                    Return typeType
                End If

                Return identifierType
            End If

            Select Case token.ParseType
                Case ParseType.Comment
                    Return commentType
                Case ParseType.Delimiter
                    Return delimiterType
                Case ParseType.Identifier
                    Return identifierType
                Case ParseType.Illegal
                    Return textType
                Case ParseType.Keyword
                    Return keywordType
                Case ParseType.NumericLiteral
                    Return numericType
                Case ParseType.Operator
                    Return operatorType
                Case ParseType.StringLiteral, ParseType.DateLiteral
                    Return stringType
                Case Else
                    Return textType
            End Select
        End Function

        Public Function GetClassificationSpans(textSpan As SnapshotSpan) As IList(Of ClassificationSpan) Implements IClassifier.GetClassificationSpans
            Dim classifications As New List(Of ClassificationSpan)
            Dim startLineNo = textSpan.Snapshot.GetLineNumberFromPosition(textSpan.Start)
            Dim endLineNo = textSpan.Snapshot.GetLineNumberFromPosition(textSpan.End)
            Dim snapshot = textSpan.Snapshot

            For i = startLineNo To endLineNo
                Dim line = snapshot.GetLineFromLineNumber(i)
                Dim tokenList = LineScanner.GetTokenEnumerator(line.GetText(), i)
                Dim token = sb.Token.Illegal
                Dim prevToken = sb.Token.Illegal
                Dim illegal = Token.Illegal

                Do
                    illegal = prevToken
                    prevToken = token
                    token = tokenList.Current
                    Dim classification = GetClassificationForType(token, prevToken, illegal)
                    Dim span = snapshot.CreateTextSpan(
                             line.Start + token.Column,
                             If(token.Text <> "", token.Text.Length, 0),
                             SpanTrackingMode.EdgeInclusive
                    )
                    classifications.Add(New ClassificationSpan(span, classification))
                Loop While tokenList.MoveNext()
            Next

            Return classifications
        End Function
    End Class
End Namespace
