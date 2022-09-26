Imports System.ComponentModel
Imports System.Windows.Markup

Namespace Microsoft.SmallVisualBasic.Shell
    <ContentProperty("CommandGroups")>
    Public Class CommandChunk
        Implements INotifyPropertyChanged

        Private nameField As String
        Private commandGroupsField As CommandGroupCollection = New CommandGroupCollection()
        Private maxSizeField As Double = Double.MaxValue

        Public Property Name As String
            Get
                Return nameField
            End Get
            Set(value As String)
                nameField = value
                Notify("Name")
            End Set
        End Property

        Public Property MaxSize As Double
            Get
                Return maxSizeField
            End Get
            Set(value As Double)
                maxSizeField = value
                Notify("MaxSize")
            End Set
        End Property

        Public ReadOnly Property CommandGroups As CommandGroupCollection
            Get
                Return commandGroupsField
            End Get
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private Sub Notify(propName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
        End Sub
    End Class
End Namespace
