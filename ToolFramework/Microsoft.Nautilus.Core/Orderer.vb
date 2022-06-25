Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports Microsoft.Contracts

Namespace Microsoft.Nautilus.Core
    <Export(GetType(IOrderer))>
    Public Class Orderer
        Implements IOrderer

        Private Class OrderNode(Of TValue)
            Private _previousNodes As New List(Of OrderNode(Of TValue))
            Private _nextNodes As New List(Of OrderNode(Of TValue))

            Public ReadOnly Property NextNodes As IList(Of OrderNode(Of TValue))
                Get
                    Return _nextNodes
                End Get
            End Property

            Public ReadOnly Property PreviousNodes As IList(Of OrderNode(Of TValue))
                Get
                    Return _previousNodes
                End Get
            End Property

            Public Property Value As TValue

            Public Property IsVisited As Boolean
        End Class

        Public Function Order(Of TValue, TMetadata As IOrderable)(
                            items As IEnumerable(Of ImportInfo(Of TValue, TMetadata))
                   ) As IEnumerable(Of ImportInfo(Of TValue, TMetadata)) _
               Implements IOrderer.Order

            Contract.RequiresNotNull(items, "items")
            Dim _orderGraph = BuildGraph(ConvertAll(items,
                        Function(valueInfo As ImportInfo(Of TValue, TMetadata)) As Collections.Generic.KeyValuePair(Of IOrderable, ComponentModel.Composition.ImportInfo(Of TValue, TMetadata))
                            Return New KeyValuePair(Of IOrderable, ImportInfo(Of TValue, TMetadata))(valueInfo.Metadata, valueInfo)
                        End Function))
            Return OrderGraph(_orderGraph)
        End Function

        Public Function Order(Of TOrderable As IOrderable)(
                        items As IEnumerable(Of TOrderable)
                    ) As IEnumerable(Of TOrderable) Implements IOrderer.Order

            Contract.RequiresNotNull(items, "items")
            Dim _orderGraph = BuildGraph(ConvertAll(items,
                     Function(orderable As TOrderable) As Collections.Generic.KeyValuePair(Of IOrderable, TOrderable)
                         Return New KeyValuePair(Of IOrderable, TOrderable)(orderable, orderable)
                     End Function))
            Return OrderGraph(_orderGraph)
        End Function

        Public Function Order(Of T)(
                           items As IEnumerable(Of T), converter1 As Converter(Of T, IOrderable)
                    ) As IEnumerable(Of T) Implements IOrderer.Order

            Contract.RequiresNotNull(items, "items")
            Contract.RequiresNotNull(converter1, "converter")
            Dim _orderGraph = BuildGraph(ConvertAll(items,
                           Function(_value As T) As Collections.Generic.KeyValuePair(Of IOrderable, T)
                               Return New KeyValuePair(Of IOrderable, T)(converter1(_value), _value)
                           End Function))
            Return OrderGraph(_orderGraph)
        End Function

        Private Iterator Function OrderGraph(Of TValue)(graph As Dictionary(Of String, OrderNode(Of TValue))) As IEnumerable(Of TValue)
            Dim sortedList = TopologicalSort(graph)
            For Each node In sortedList
                If node.Value IsNot Nothing Then
                    Yield node.Value
                End If
            Next
        End Function

        Private Iterator Function ConvertAll(Of TSource, TTarget)(
                        sourceList As IEnumerable(Of TSource),
                        converter As Converter(Of TSource, TTarget)
                    ) As IEnumerable(Of TTarget)

            For Each source As TSource In sourceList
                Yield converter(source)
            Next
        End Function

        Private Function TopologicalSort(Of TValue)(orderGraph As Dictionary(Of String, OrderNode(Of TValue))) As List(Of OrderNode(Of TValue))
            Dim list1 As New List(Of OrderNode(Of TValue))
            Dim queue1 As New Queue(Of OrderNode(Of TValue))
            For Each _value As OrderNode(Of TValue) In orderGraph.Values
                If Not _value.IsVisited Then
                    VisitNode(_value, queue1)
                End If
                While queue1.Count > 0
                    list1.Insert(0, queue1.Dequeue())
                End While
            Next
            Return list1
        End Function

        Private Sub VisitNode(Of TValue)(node As OrderNode(Of TValue), orderQueue As Queue(Of OrderNode(Of TValue)))
            node.IsVisited = True
            For Each nextNode As OrderNode(Of TValue) In node.NextNodes
                If Not nextNode.IsVisited Then
                    VisitNode(nextNode, orderQueue)
                End If
            Next

            orderQueue.Enqueue(node)
        End Sub

        Private Function BuildGraph(Of TValue)(items As IEnumerable(Of KeyValuePair(Of IOrderable, TValue))) As Dictionary(Of String, OrderNode(Of TValue))
            Dim dictionary1 As New Dictionary(Of String, OrderNode(Of TValue))
            For Each item As KeyValuePair(Of IOrderable, TValue) In items
                Dim orCreateNode As OrderNode(Of TValue) = GetOrCreateNode(dictionary1, item.Key.Name)
                orCreateNode.Value = item.Value
                If item.Key.Before IsNot Nothing Then
                    For Each item2 As String In item.Key.Before
                        Dim orCreateNode2 As OrderNode(Of TValue) = GetOrCreateNode(dictionary1, item2)
                        If Not orCreateNode.NextNodes.Contains(orCreateNode2) Then
                            orCreateNode.NextNodes.Add(orCreateNode2)
                            orCreateNode2.PreviousNodes.Add(orCreateNode)
                        End If
                    Next
                End If

                If item.Key.After Is Nothing Then
                    Continue For
                End If

                For Each item3 As String In item.Key.After
                    Dim orCreateNode3 As OrderNode(Of TValue) = GetOrCreateNode(dictionary1, item3)
                    If Not orCreateNode3.NextNodes.Contains(orCreateNode) Then
                        orCreateNode3.NextNodes.Add(orCreateNode)
                        orCreateNode.PreviousNodes.Add(orCreateNode3)
                    End If
                Next
            Next
            Return dictionary1
        End Function

        Private Function GetOrCreateNode(Of TValue)(orderGraph As Dictionary(Of String, OrderNode(Of TValue)), name1 As String) As OrderNode(Of TValue)
            Contract.Requires(If((name1 IsNot Nothing), Nothing, New InvalidOperationException("Orderable elements must all have a non-null Name.")))
            Dim _value As Microsoft.Nautilus.Core.Orderer.OrderNode(Of TValue) = Nothing
            If Not orderGraph.TryGetValue(name1, _value) Then
                _value = New OrderNode(Of TValue)
                orderGraph(name1) = _value
            End If

            Return _value
        End Function

    End Class
End Namespace
