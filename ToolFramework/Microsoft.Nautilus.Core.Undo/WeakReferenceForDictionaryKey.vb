
Namespace Microsoft.Nautilus.Core.Undo
    Friend Class WeakReferenceForDictionaryKey
        Inherits WeakReference

        Private Const xorWeakReference As Integer = 52428
        Private ReadOnly hashCode As Integer

        Public Sub New(target As Object)
            MyBase.New(target)
            hashCode = target.GetHashCode() Xor &HCCCC
        End Sub

        Public Overrides Function GetHashCode() As Integer
            Return hashCode
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Dim flag As Boolean = False
            Dim weakReferenceForDictionaryKey1 As WeakReferenceForDictionaryKey = TryCast(obj, WeakReferenceForDictionaryKey)

            If Object.ReferenceEquals(weakReferenceForDictionaryKey1, Nothing) Then
                Return False
            End If

            If Object.ReferenceEquals(Me, weakReferenceForDictionaryKey1) Then
                Return True
            End If

            Dim obj2 As Object = Nothing
            Dim obj3 As Object = Nothing
            Try
                obj2 = Target
                obj3 = weakReferenceForDictionaryKey1.Target
            Catch
            End Try

            If obj2 Is Nothing OrElse obj3 Is Nothing Then
                Return False
            End If

            Return Object.Equals(obj2, obj3)
        End Function
    End Class
End Namespace
