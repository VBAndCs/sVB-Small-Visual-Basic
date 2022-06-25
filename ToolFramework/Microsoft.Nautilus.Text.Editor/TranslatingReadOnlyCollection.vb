Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Namespace Microsoft.Nautilus.Text.Editor
    Friend NotInheritable Class TranslatingReadOnlyCollection(Of Base As Class, Derived As {Class, Base})
        Implements IList(Of Base), ICollection(Of Base), IEnumerable(Of Base), IEnumerable

        Private NotInheritable Class MyEnumerator
            Implements IEnumerator(Of Base), IDisposable, IEnumerator

            Private _currentIndex As Integer = -1
            Private _source As ReadOnlyCollection(Of Derived)

            Public ReadOnly Property Current As Base Implements IEnumerator(Of Base).Current
                Get
                    Return CType(_source(_currentIndex), Base)
                End Get
            End Property

            Public ReadOnly Property IEnumerator_Current As Object Implements IEnumerator.Current
                Get
                    Return _source(_currentIndex)
                End Get
            End Property

            Public Sub New(source As ReadOnlyCollection(Of Derived))
                _source = source
            End Sub

            Private Sub Dispose() Implements IDisposable.Dispose
            End Sub

            Public Function MoveNext() As Boolean Implements Collections.IEnumerator.MoveNext
                _currentIndex += 1
                Return _currentIndex < _source.Count
            End Function

            Public Sub Reset() Implements Collections.IEnumerator.Reset
                _currentIndex = -1
            End Sub
        End Class

        Private _source As ReadOnlyCollection(Of Derived)

        Default Public Property Item(index As Integer) As Base Implements Collections.Generic.IList(Of Base).Item
            Get
                Return CType(_source(index), Base)
            End Get

            Set(value As Base)
                Throw New InvalidOperationException("Attempt to modify a read only collection.")
            End Set
        End Property

        Public ReadOnly Property Count As Integer Implements ICollection(Of Base).Count
            Get
                Return _source.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean = True Implements ICollection(Of Base).IsReadOnly

        Public Sub New(source As ReadOnlyCollection(Of Derived))
            _source = source
        End Sub

        Public Function IndexOf(item As Base) As Integer Implements Collections.Generic.IList(Of Base).IndexOf
            Dim val As Derived = TryCast(item, Derived)

            If val IsNot Nothing OrElse item Is Nothing Then
                Return _source.IndexOf(val)
            End If

            Return -1
        End Function

        Private Sub Insert(index As Integer, item As Base) Implements Collections.Generic.IList(Of Base).Insert
            Throw New InvalidOperationException("Attempt to modify a read only collection.")
        End Sub

        Private Sub RemoveAt(index As Integer) Implements Collections.Generic.IList(Of Base).RemoveAt
            Throw New InvalidOperationException("Attempt to modify a read only collection.")
        End Sub

        Private Sub Add(item As Base) Implements Collections.Generic.ICollection(Of Base).Add
            Throw New InvalidOperationException("Attempt to modify a read only collection.")
        End Sub

        Private Sub Clear() Implements Collections.Generic.ICollection(Of Base).Clear
            Throw New InvalidOperationException("Attempt to modify a read only collection.")
        End Sub

        Public Function Contains(item As Base) As Boolean Implements Collections.Generic.ICollection(Of Base).Contains
            Dim val As Derived = TryCast(item, Derived)

            If val IsNot Nothing OrElse item Is Nothing Then
                Return _source.Contains(val)
            End If

            Return False
        End Function

        Public Sub CopyTo(array As Base(), arrayIndex As Integer) Implements Collections.Generic.ICollection(Of Base).CopyTo
            For i As Integer = 0 To _source.Count - 1
                array(i + arrayIndex) = CType(_source(i), Base)
            Next
        End Sub

        Private Function Remove(item As Base) As Boolean Implements Collections.Generic.ICollection(Of Base).Remove
            Throw New InvalidOperationException("Attempt to modify a read only collection.")
        End Function

        Public Function GetEnumerator() As IEnumerator(Of Base) Implements Collections.Generic.IEnumerable(Of Base).GetEnumerator
            Return New MyEnumerator(_source)
        End Function

        Private Function GetEnumerator2() As IEnumerator Implements Collections.IEnumerable.GetEnumerator
            Return New MyEnumerator(_source)
        End Function
    End Class
End Namespace
