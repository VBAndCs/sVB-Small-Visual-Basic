Imports System.ComponentModel
Imports System.Windows.Markup

Namespace Microsoft.SmallVisualBasic.Shell
    <ContentProperty("CommandChunks")>
    Public Class RibbonTab
        Implements INotifyPropertyChanged

        Private _name As String
        Private _commandChunks As New CommandChunkCollection()

        Public ReadOnly Property CommandChunks As CommandChunkCollection
            Get
                Return _commandChunks
            End Get
        End Property

        Public Property Name As String
            Get
                Return _name
            End Get

            Set(value As String)
                _name = value
                Notify("Name")
            End Set
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private Sub Notify(propName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
        End Sub
    End Class
End Namespace
