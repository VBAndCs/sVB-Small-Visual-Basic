Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Classification

Namespace Microsoft.SmallBasic.LanguageService
    Public NotInheritable Class SmallBasicClassifierProvider
        <Import>
        Public Property ClassificationTypeRegistry As IClassificationTypeRegistry

        <ContentType("text.smallbasic")>
        <Export(GetType(ClassifierProvider))>
        Public Function CreateClassifier(textBuffer As ITextBuffer) As IClassifier
            Return New SmallBasicClassifier(textBuffer, ClassificationTypeRegistry)
        End Function
    End Class
End Namespace
