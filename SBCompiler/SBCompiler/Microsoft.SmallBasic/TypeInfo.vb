Imports System
Imports System.Collections.Generic
Imports System.Reflection

Namespace Microsoft.SmallBasic
    <Serializable>
    Public Class TypeInfo
        Public Methods As New Dictionary(Of String, MethodInfo)
        Public Properties As New Dictionary(Of String, PropertyInfo)
        Public Events As New Dictionary(Of String, EventInfo)
        Public Property Type As Type
        Public Property HideFromIntellisense As Boolean
    End Class
End Namespace
