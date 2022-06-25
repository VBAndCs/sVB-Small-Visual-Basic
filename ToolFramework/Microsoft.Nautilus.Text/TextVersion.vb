
Namespace Microsoft.Nautilus.Text
    Public MustInherit Class TextVersion

        Public MustOverride ReadOnly Property [Next] As TextVersion

        Public MustOverride ReadOnly Property Changes As INormalizedTextChangeCollection

        Public ReadOnly Property TextBuffer As ITextBuffer

        Protected ReadOnly Property VersionNumber As Integer


        Protected Sub New(textBuffer As ITextBuffer, versionNumber As Integer)
            _TextBuffer = textBuffer
            _VersionNumber = versionNumber
        End Sub

        Public Shared Operator <(v1 As TextVersion, v2 As TextVersion) As Boolean
            If v1 Is Nothing Then
                Throw New ArgumentNullException("v1")
            End If

            If v2 Is Nothing Then
                Throw New ArgumentNullException("v2")
            End If

            If v1.TextBuffer IsNot v2.TextBuffer Then
                Throw New ArgumentException("The TextVersions are not for the same TextBuffer")
            End If

            Return v1._VersionNumber < v2._VersionNumber
        End Operator

        Public Shared Operator <=(v1 As TextVersion, v2 As TextVersion) As Boolean
            If v1 Is Nothing Then
                Throw New ArgumentNullException("v1")
            End If
            If v2 Is Nothing Then
                Throw New ArgumentNullException("v2")
            End If

            If v1.TextBuffer IsNot v2.TextBuffer Then
                Throw New ArgumentException("The TextVersions are not for the same TextBuffer")
            End If

            Return v1._VersionNumber <= v2._VersionNumber
        End Operator

        Public Shared Operator >(v1 As TextVersion, v2 As TextVersion) As Boolean
            Return Not (v1 <= v2)
        End Operator

        Public Shared Operator >=(v1 As TextVersion, v2 As TextVersion) As Boolean
            Return Not (v1 < v2)
        End Operator

        Public Shared Function Compare(v1 As TextVersion, v2 As TextVersion) As Integer
            If v1 Is Nothing Then
                Throw New ArgumentNullException("v1")
            End If

            If v2 Is Nothing Then
                Throw New ArgumentNullException("v2")
            End If

            If v1.TextBuffer IsNot v2.TextBuffer Then
                Throw New ArgumentException("The TextVersions are not for the same TextBuffer")
            End If

            Return v1._VersionNumber - v2._VersionNumber
        End Function
    End Class
End Namespace
