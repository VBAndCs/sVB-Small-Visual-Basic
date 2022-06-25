Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports System.Globalization
Imports Microsoft.Nautilus.Core.Undo.DataExports
Imports System.Runtime.InteropServices

Namespace Microsoft.Nautilus.Core.Undo
    <Export(GetType(IUndoHistoryRegistry))>
    Public Class UndoHistoryRegistryImpl
        Implements IUndoHistoryRegistry

        Private _histories As Dictionary(Of UndoHistory, Integer)
        Private weakContextMapping As Dictionary(Of WeakReferenceForDictionaryKey, UndoHistory)
        Private strongContextMapping As Dictionary(Of Object, UndoHistory)
        Private markers As Dictionary(Of String, UndoTransactionMarker)
        Private linkManager As LinkedUndoTransactionManager
        Private _undoTransactionMarkers As IEnumerable(Of UndoTransactionMarker)

        <Import("UndoTransactionMarker")>
        Public Property UndoTransactionMarkers As IEnumerable(Of UndoTransactionMarker) Implements IUndoHistoryRegistry.UndoTransactionMarkers
            Get
                Return _undoTransactionMarkers
            End Get

            Set(value As IEnumerable(Of UndoTransactionMarker))
                _undoTransactionMarkers = value
                LoadUndoTransactionMarkers()
            End Set
        End Property

        Public ReadOnly Property Histories As IEnumerable(Of UndoHistory) Implements IUndoHistoryRegistry.Histories
            Get
                Return _histories.Keys
            End Get
        End Property

        Public ReadOnly Property GlobalHistory As UndoHistory Implements IUndoHistoryRegistry.GlobalHistory

        Friend ReadOnly Property LinkedUndoTransactionManager As LinkedUndoTransactionManager
            Get
                Return linkManager
            End Get
        End Property

        Public Sub New()
            _histories = New Dictionary(Of UndoHistory, Integer)
            weakContextMapping = New Dictionary(Of WeakReferenceForDictionaryKey, UndoHistory)
            strongContextMapping = New Dictionary(Of Object, UndoHistory)
            _GlobalHistory = New UndoHistoryImpl(Me)
            _histories.Add(_GlobalHistory, 1)
            markers = New Dictionary(Of String, UndoTransactionMarker)
            linkManager = New LinkedUndoTransactionManager
        End Sub

        Public Function RegisterHistory(context As Object) As UndoHistory Implements IUndoHistoryRegistry.RegisterHistory
            If context Is Nothing Then
                Throw New ArgumentNullException("context", String.Format(CultureInfo.CurrentCulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"RegisterHistory", "context"}))
            End If
            Return RegisterHistory(context, keepAlive:=False)
        End Function

        Public Function RegisterHistory(context As Object, keepAlive As Boolean) As UndoHistory Implements IUndoHistoryRegistry.RegisterHistory
            If context Is Nothing Then
                Throw New ArgumentNullException("context", String.Format(CultureInfo.CurrentCulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"RegisterHistory", "context"}))
            End If
            Dim undoHistory1 As UndoHistory
            If strongContextMapping.ContainsKey(context) Then
                undoHistory1 = strongContextMapping(context)
                If Not keepAlive Then
                    strongContextMapping.Remove(context)
                    weakContextMapping.Add(New WeakReferenceForDictionaryKey(context), undoHistory1)
                End If

            ElseIf weakContextMapping.ContainsKey(New WeakReferenceForDictionaryKey(context)) Then
                undoHistory1 = weakContextMapping(New WeakReferenceForDictionaryKey(context))
                If keepAlive Then
                    weakContextMapping.Remove(New WeakReferenceForDictionaryKey(context))
                    strongContextMapping.Add(context, undoHistory1)
                End If

            Else
                undoHistory1 = New UndoHistoryImpl(Me)
                _histories.Add(undoHistory1, 1)
                If keepAlive Then
                    strongContextMapping.Add(context, undoHistory1)
                Else
                    weakContextMapping.Add(New WeakReferenceForDictionaryKey(context), undoHistory1)
                End If
            End If

            Return undoHistory1
        End Function

        Public Function GetHistory(context As Object) As UndoHistory Implements IUndoHistoryRegistry.GetHistory
            If context Is Nothing Then
                Throw New ArgumentNullException("context", String.Format(CultureInfo.CurrentCulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"GetHistory", "context"}))
            End If

            If strongContextMapping.ContainsKey(context) Then
                Return strongContextMapping(context)
            End If

            If weakContextMapping.ContainsKey(New WeakReferenceForDictionaryKey(context)) Then
                Return weakContextMapping(New WeakReferenceForDictionaryKey(context))
            End If

            Throw New InvalidOperationException("The request to retrieve an UndoHistory cannot be satisfied, because no UndoHistory has been associated with the given context.")
        End Function

        Public Function TryGetHistory(context As Object, <Out> ByRef history As UndoHistory) As Boolean Implements IUndoHistoryRegistry.TryGetHistory
            If context Is Nothing Then
                Throw New ArgumentNullException("context", String.Format(CultureInfo.CurrentCulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"TryGetHistory", "context"}))
            End If

            Dim undoHistory1 As UndoHistory = Nothing
            If strongContextMapping.ContainsKey(context) Then
                undoHistory1 = strongContextMapping(context)
            ElseIf weakContextMapping.ContainsKey(New WeakReferenceForDictionaryKey(context)) Then
                undoHistory1 = weakContextMapping(New WeakReferenceForDictionaryKey(context))
            End If

            history = undoHistory1
            Return undoHistory1 IsNot Nothing
        End Function

        Public Sub AttachHistory(context As Object, history As UndoHistory) Implements IUndoHistoryRegistry.AttachHistory
            If context Is Nothing Then
                Throw New ArgumentNullException("context", String.Format(CultureInfo.CurrentCulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"AttachHistory", "context"}))
            End If

            If history Is Nothing Then
                Throw New ArgumentNullException("context", String.Format(CultureInfo.CurrentCulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"AttachHistory", "history"}))
            End If
            AttachHistory(context, history, keepAlive:=False)
        End Sub

        Public Sub AttachHistory(context As Object, history As UndoHistory, keepAlive As Boolean) Implements IUndoHistoryRegistry.AttachHistory
            If context Is Nothing Then
                Throw New ArgumentNullException("context", String.Format(CultureInfo.CurrentCulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"AttachHistory", "context"}))
            End If

            If history Is Nothing Then
                Throw New ArgumentNullException("context", String.Format(CultureInfo.CurrentCulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"AttachHistory", "history"}))
            End If

            If strongContextMapping.ContainsKey(context) OrElse weakContextMapping.ContainsKey(New WeakReferenceForDictionaryKey(context)) Then
                Throw New InvalidOperationException("The request to attach an UndoHistory with the given context cannot be satisfied, because that context is already attached to a History in the registry.")
            End If

            If Not _histories.ContainsKey(history) Then
                _histories.Add(history, 1)
            Else
                _histories(history) += 1
            End If

            If keepAlive Then
                strongContextMapping.Add(context, history)
            Else
                weakContextMapping.Add(New WeakReferenceForDictionaryKey(context), history)
            End If
        End Sub

        Public Sub DetachHistory(context As Object) Implements IUndoHistoryRegistry.DetachHistory
            If context Is Nothing Then
                Throw New ArgumentNullException("context", String.Format(CultureInfo.CurrentCulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"DetachHistory", "context"}))
            End If

            Dim undoHistory1 As UndoHistory = Nothing
            If strongContextMapping.ContainsKey(context) Then
                undoHistory1 = strongContextMapping(context)
                strongContextMapping.Remove(context)
            ElseIf weakContextMapping.ContainsKey(New WeakReferenceForDictionaryKey(context)) Then
                undoHistory1 = weakContextMapping(New WeakReferenceForDictionaryKey(context))
                weakContextMapping.Remove(New WeakReferenceForDictionaryKey(context))
            End If

            If undoHistory1 IsNot Nothing Then
                _histories(undoHistory1) -= 1
                If _histories(undoHistory1) <= 0 Then
                    _histories.Remove(undoHistory1)
                End If
            End If
        End Sub

        Public Sub RemoveHistory(history As UndoHistory) Implements IUndoHistoryRegistry.RemoveHistory
            If history Is Nothing Then
                Throw New ArgumentNullException("context", String.Format(CultureInfo.CurrentCulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"RemoveHistory", "history"}))
            End If

            If Object.ReferenceEquals(history, _GlobalHistory) Then
                Throw New ArgumentException("The request to remove the given history from the registry cannot be satisfied, because the given history is the GlobalHistory.", "history")
            End If

            If Not _histories.ContainsKey(history) Then
                Return
            End If

            _histories.Remove(history)
            Dim list1 As New List(Of Object)
            For Each key As Object In strongContextMapping.Keys
                If Object.ReferenceEquals(strongContextMapping(key), history) Then
                    list1.Add(key)
                End If
            Next

            list1.ForEach(Sub(o As Object)
                              strongContextMapping.Remove(o)
                          End Sub)
            Dim list2 As New List(Of WeakReferenceForDictionaryKey)
            For Each key2 As WeakReferenceForDictionaryKey In weakContextMapping.Keys
                If Object.ReferenceEquals(weakContextMapping(key2), history) Then
                    list2.Add(key2)
                End If
            Next

            list2.ForEach(Sub(o As WeakReferenceForDictionaryKey)
                              weakContextMapping.Remove(o)
                          End Sub)
        End Sub

        Public Function GetTotalHistorySize() As Long Implements IUndoHistoryRegistry.GetTotalHistorySize
            Dim num As Long = 0L
            For Each key As UndoHistory In _histories.Keys
                num += key.GetEstimatedSize()
            Next
            Return num
        End Function

        Protected Sub LoadUndoTransactionMarkers()
            markers.Clear()
            For Each undoTransactionMarker1 As UndoTransactionMarker In UndoTransactionMarkers
                markers.Add(undoTransactionMarker1.Name, undoTransactionMarker1)
            Next
        End Sub

        Public Function GetUndoTransactionMarker(name1 As String) As UndoTransactionMarker Implements IUndoHistoryRegistry.GetUndoTransactionMarker
            If String.IsNullOrEmpty(name1) Then
                Throw New ArgumentNullException("name", String.Format(CultureInfo.CurrentCulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"GetUndoTransactionMarker", "name"}))
            End If
            Dim value As UndoTransactionMarker = Nothing
            If Not markers.TryGetValue(name1, value) Then
                Throw New InvalidOperationException(String.Format(CultureInfo.CurrentCulture, "There is no UndoTransactionMarker with name '{0}'.", New Object(0) {name1}))
            End If
            Return value
        End Function

        Public Sub LimitHistoryDepth(depth As Integer) Implements IUndoHistoryRegistry.LimitHistoryDepth
            Throw New NotImplementedException("The method or operation is not implemented.")
        End Sub

        Public Sub LimitHistorySize(size As Long) Implements IUndoHistoryRegistry.LimitHistorySize
            Throw New NotImplementedException("The method or operation is not implemented.")
        End Sub

        Public Sub LimitTotalHistorySize(size As Long) Implements IUndoHistoryRegistry.LimitTotalHistorySize
            Throw New NotImplementedException("The method or operation is not implemented.")
        End Sub

        Public Function CreateLinkedUndoTransaction(description As String) As LinkedUndoTransaction Implements IUndoHistoryRegistry.CreateLinkedUndoTransaction
            If String.IsNullOrEmpty(description) Then
                Throw New ArgumentNullException("description", String.Format(CultureInfo.CurrentUICulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"CreateTransaction", "description"}))
            End If
            Return linkManager.CreateLinkedUndoTransaction(description)
        End Function
    End Class
End Namespace
