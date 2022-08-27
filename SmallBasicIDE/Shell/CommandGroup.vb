Imports System.Windows.Markup

Namespace Microsoft.SmallBasic.Shell
    <ContentProperty("Commands")>
    Public MustInherit Class CommandGroup
        Private nameField As String
        Private commandsField As CommandRegistry = New CommandRegistry()
        Private maxSizeField As Double = Double.MaxValue

        Public Property MaxSize As Double
            Get
                Return maxSizeField
            End Get
            Set(value As Double)
                maxSizeField = value
            End Set
        End Property

        Public Property Name As String
            Get
                Return nameField
            End Get
            Set(value As String)
                nameField = value
            End Set
        End Property

        Public ReadOnly Property Commands As CommandRegistry
            Get
                Return commandsField
            End Get
        End Property

        Public Sub New()
        End Sub
    End Class
End Namespace
