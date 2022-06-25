Imports System.Collections.Generic
Imports System.Runtime.Serialization
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Classification
    <Serializable>
    Friend NotInheritable Class AmbientClassificationFormatMap
        Private _properties As Dictionary(Of String, TextFormattingRunProperties)
        Public ReadOnly Property Properties As Dictionary(Of String, TextFormattingRunProperties)
            Get
                Return _properties
            End Get
        End Property

        Friend Sub New()
            _properties = New Dictionary(Of String, TextFormattingRunProperties)
        End Sub
        Friend Sub Deserializing(context As StreamingContext)
            OnDeserialized(context)
        End Sub
        <OnDeserialized>
        Private Sub OnDeserialized(context As StreamingContext)
        End Sub
        Public Sub Update()
        End Sub
    End Class
End Namespace
