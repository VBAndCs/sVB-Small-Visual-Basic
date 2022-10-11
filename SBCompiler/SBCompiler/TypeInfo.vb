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

        Dim _name As String
        Public Property Name As String
            Get
                Return If(_name, Type.Name)
            End Get

            Set(value As String)
                _name = value
            End Set
        End Property
    End Class
End Namespace
