Imports System.ComponentModel
Imports System.Windows.Markup

Namespace Microsoft.SmallVisualBasic.Shell
    <ContentProperty("CommandGroups")>
    Public Class CommandChunk
        Implements INotifyPropertyChanged

        Private _name As String
        Private _commandGroups As CommandGroupCollection = New CommandGroupCollection()
        Private _maxSize As Double = Double.MaxValue

        Public Property Name As String
            Get
                Return _name
            End Get

            Set(value As String)
                _name = value
                Notify("Name")
            End Set
        End Property

        Public Property MaxSize As Double
            Get
                Return _maxSize
            End Get

            Set(value As Double)
                _maxSize = value
                Notify("MaxSize")
            End Set
        End Property

        Public ReadOnly Property CommandGroups As CommandGroupCollection
            Get
                Return _commandGroups
            End Get
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private Sub Notify(propName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
        End Sub
    End Class
End Namespace
