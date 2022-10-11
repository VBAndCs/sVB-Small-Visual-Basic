Imports System
Imports System.Collections.Generic
Imports System.Reflection

Namespace Microsoft.SmallVisualBasic
    <Serializable>
    Public Class TypeInfo
        Public Methods As New Dictionary(Of String, MethodInfo)
        Public Properties As New Dictionary(Of String, PropertyInfo)
        Public Events As New Dictionary(Of String, EventInfo)
        Public Property Type As Type
        Public Property HideFromIntellisense As Boolean

        Public ReadOnly Property Key As String
            Get
                Return If(_Name, "").ToLower()
            End Get
        End Property

        Public Property Name As String
    End Class
End Namespace
