Namespace Microsoft.Nautilus.Text.StringRebuilder
    Friend Interface ITreeNode
        ReadOnly Property Depth As Integer

        ReadOnly Property StringRebuilder As IStringRebuilder

        Function Child(side1 As Side) As ITreeNode

        Function OtherChild(side1 As Side) As ITreeNode
    End Interface
End Namespace
