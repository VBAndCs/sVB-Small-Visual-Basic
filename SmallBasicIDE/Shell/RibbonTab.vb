Imports System.ComponentModel
Imports System.Windows.Markup

Namespace Microsoft.SmallBasic.Shell
    <ContentProperty("CommandChunks")>
    Public Class RibbonTab
        Implements INotifyPropertyChanged

        Private nameField As String
        Private commandChunksField As CommandChunkCollection = New CommandChunkCollection()

        Public ReadOnly Property CommandChunks As CommandChunkCollection
            Get
                Return commandChunksField
            End Get
        End Property

        Public Property Name As String
            Get
                Return nameField
            End Get
            Set(ByVal value As String)
                nameField = value
                Notify("Name")
            End Set
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private Sub Notify(ByVal propName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
        End Sub
    End Class
End Namespace
