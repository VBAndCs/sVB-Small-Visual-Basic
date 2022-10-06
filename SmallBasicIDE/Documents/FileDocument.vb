Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.IO
Imports Microsoft.SmallBasic

Namespace Microsoft.SmallVisualBasic.Documents
    Public Class FileDocument
        Implements INotifyPropertyChanged

        Protected _file As String
        Private _isDirty As Boolean
        Private _propertyStore As New Dictionary(Of Object, Object)()

        Public Property File As String
            Get
                Return _file
            End Get

            Friend Set(value As String)
                _file = value
                NotifyProperty("FilePath")
                NotifyProperty("Title")
            End Set
        End Property

        Public Property IsDirty As Boolean
            Get
                Return _isDirty
            End Get

            Friend Set(value As Boolean)
                If _isDirty <> value Then
                    _isDirty = value
                    NotifyProperty("IsDirty")
                    NotifyProperty("Title")
                End If
            End Set
        End Property

        Public Property IsNew As Boolean

        Public ReadOnly Property PropertyStore As Dictionary(Of Object, Object)
            Get
                Return _propertyStore
            End Get
        End Property

        Public Overridable ReadOnly Property Title As String
            Get
                Dim text = If(IsDirty, " *", "")

                If IsNew Then
                    Return ResourceHelper.GetString("Untitled") & text
                End If

                Return $"{Path.GetFileName(_file)}{text} - {Path.GetFullPath(_file)}"
            End Get
        End Property

        Public Event Closed As EventHandler
        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Public Sub New(filePath As String)
            If filePath <> "" AndAlso Not IO.File.Exists(filePath) Then
                IO.File.WriteAllText(filePath, "")
            End If


            If filePath <> "" AndAlso Not Path.IsPathRooted(filePath) Then
                _file = Path.GetFullPath(filePath)
            Else
                _file = filePath
            End If
        End Sub

        Public Function Open() As Stream
            Return IO.File.Open(_file, FileMode.OpenOrCreate, FileAccess.ReadWrite)
        End Function

        Public Overridable Sub Close()
            RaiseEvent Closed(Me, EventArgs.Empty)
        End Sub

        Protected Sub NotifyProperty(propertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
        End Sub
    End Class
End Namespace
