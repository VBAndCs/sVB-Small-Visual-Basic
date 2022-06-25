Imports System.Windows
Imports System.Windows.Media.TextFormatting

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class TextFormattingModifier
        Inherits TextModifier

        Public Overrides ReadOnly Property FlowDirection As FlowDirection
            Get
                Return FlowDirection.LeftToRight
            End Get
        End Property

        Public Overrides ReadOnly Property HasDirectionalEmbedding As Boolean = True

        Public Overrides ReadOnly Property Length As Integer = 1

        Public Overrides ReadOnly Property Properties As TextRunProperties

        Public Sub New(properties As TextRunProperties)
            _Properties = properties
        End Sub

        Public Overrides Function ModifyProperties(properties As TextRunProperties) As TextRunProperties
            Return _Properties
        End Function
    End Class
End Namespace
