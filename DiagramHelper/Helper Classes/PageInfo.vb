Public Structure PageInfo
    Public Name As String
    Public IsNew As Boolean
    Public HasChanges As Boolean
    Public Xaml As String
    Public DocPath As String
    Public XamlPath As String

    Public Sub New(name As String,
                   isNew As Boolean,
                   hasChanges As Boolean,
                   xaml As String,
                   docPath As String,
                   xamlPath As String)

        Me.Name = name
        Me.Xaml = xaml
        Me.HasChanges = hasChanges
        Me.IsNew = isNew
        Me.DocPath = docPath
        Me.XamlPath = xamlPath
    End Sub

    Public Shared Widening Operator CType(
               info As (Name As String, IsNew As Boolean, HasChanges As Boolean, Xaml As String, DocPath As String, XamlPath As String)
            ) As PageInfo

        Return New PageInfo(info.Name, info.IsNew, info.HasChanges, info.Xaml, info.DocPath, info.XamlPath)
    End Operator



End Structure

