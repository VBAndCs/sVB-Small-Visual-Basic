Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication

Namespace WinForms

    Friend NotInheritable Class ListBase
        Private Shared Function GetSelector(listName As String) As Wpf.Primitives.Selector
            Dim c = Control.GetControl(listName)
            Dim lst = TryCast(c, Wpf.Primitives.Selector)

            If lst Is Nothing Then
                Throw New Exception($"{listName} is not a name of a ListBox or a ComboBox control.")
            End If
            Return lst
        End Function

        Friend Shared Function GetItemsCount(listName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetItemsCount = GetSelector(listName).Items.Count
                    Catch ex As Exception
                        Control.ReportError(listName, "ItemsCount", ex)
                    End Try
                End Sub)
        End Function

        Friend Shared Function GetItems(listName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim map = New Dictionary(Of Primitive, Primitive)
                        Dim num = 1
                        For Each item As String In GetSelector(listName).Items
                            map(num) = item
                            num += 1
                        Next
                        GetItems = Primitive.ConvertFromMap(map)

                    Catch ex As Exception
                        Control.ReportError(listName, "ItemsCount", ex)
                    End Try
                End Sub)
        End Function

        Friend Shared Function GetSelectedItem(listName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim item As String = GetSelector(listName).SelectedItem
                        GetSelectedItem = item
                    Catch ex As Exception
                        Control.ReportError(listName, "SelectedItem", ex)
                    End Try
                End Sub)
        End Function

        Friend Shared Sub SetSelectedItem(listName As Primitive, item As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetSelector(listName).SelectedItem = CStr(item)
                    Catch ex As Exception
                        Control.ReportError(listName, "SelectedItem", item, ex)
                    End Try
                End Sub)
        End Sub

        Friend Shared Function GetSelectedIndex(listName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetSelectedIndex = GetSelector(listName).SelectedIndex + 1
                    Catch ex As Exception
                        Control.ReportError(listName, "SelectedIndex", ex)
                    End Try
                End Sub)
        End Function

        Friend Shared Sub SetSelectedIndex(listName As Primitive, index As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetSelector(listName)
                        Dim i As Integer
                        If Not index.IsNumber Then
                            i = -1
                        Else
                            i = index
                        End If

                        If i < 1 OrElse i > lst.Items.Count Then
                            i = -1
                        Else
                            i -= 1
                        End If

                        lst.SelectedIndex = i

                    Catch ex As Exception
                        Control.ReportError(listName, "SelectedIndex", index, ex)
                    End Try
                End Sub)
        End Sub

        Friend Shared Function GetItemAt(listName As Primitive, index As Primitive) As Primitive
            App.Invoke(
                Sub()

                    Try
                        Dim lst = GetSelector(listName)
                        If Not index.IsNumber Then
                            GetItemAt = ""
                        Else
                            Dim i As Integer = index
                            If i < 1 OrElse i > lst.Items.Count Then
                                GetItemAt = ""
                            Else
                                GetItemAt = lst.Items(i - 1).ToString()
                            End If
                        End If
                    Catch ex As Exception
                        Control.ReportSubError(listName, "GetItemAt", ex)
                    End Try
                End Sub)
        End Function

        Friend Shared Function SetItemAt(
                          listName As Primitive,
                          index As Primitive,
                          value As Primitive
                   ) As Boolean

            App.Invoke(
                Sub()
                    Try
                        If Not index.IsNumber Then
                            SetItemAt = False
                            Return
                        End If

                        Dim lst = GetSelector(listName)
                        Dim i As Integer = index

                        If i > 0 AndAlso i <= lst.Items.Count Then
                            lst.Items(i - 1) = CStr(value)
                            SetItemAt = True
                        Else
                            SetItemAt = False
                        End If

                    Catch ex As Exception
                        SetItemAt = False
                        Control.ReportSubError(listName, "SetItemAt", ex)
                    End Try
                End Sub)
        End Function

        Friend Shared Function AddItem(listName As Primitive, value As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetSelector(listName)
                        Dim id As Integer

                        If value.IsArray Then
                            Dim items = value._arrayMap
                            If items Is Nothing OrElse items.Count = 0 Then
                                AddItem = 0
                                Return
                            End If

                            For Each item In items.Values
                                id = lst.Items.Add(CStr(item))
                            Next
                        Else
                            id = lst.Items.Add(CStr(value))
                        End If

                        AddItem = id + 1

                    Catch ex As Exception
                        AddItem = 0
                        Control.ReportSubError(listName, "AddItem", ex)
                    End Try
                End Sub)
        End Function

        Friend Shared Function AddItemAt(listName As Primitive, value As Primitive, index As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        If Not index.IsNumber Then Return

                        Dim i As Integer = CInt(index) - 1
                        Dim lst = GetSelector(listName)
                        If i < 0 OrElse i > lst.Items.Count Then
                            AddItemAt = False
                            Return
                        End If

                        If value.IsArray Then
                            Dim items = value._arrayMap.Values
                            If items Is Nothing OrElse items.Count = 0 Then
                                AddItemAt = False
                                Return
                            End If

                            For n = items.Count - 1 To 0 Step -1
                                lst.Items.Insert(i, CStr(items(n)))
                            Next
                        Else
                            lst.Items.Insert(i, value.AsString())
                        End If

                        AddItemAt = True

                    Catch ex As Exception
                        AddItemAt = False
                        Control.ReportSubError(listName, "AddItemAt", ex)
                    End Try

                End Sub)
        End Function

        Friend Shared Function RemoveItem(listName As Primitive, value As Primitive) As Boolean
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetSelector(listName)
                        If value.IsArray Then
                            RemoveItem = False
                            For Each value In value._arrayMap.Values
                                If RemoveItem(value, lst) Then RemoveItem = True
                            Next
                        Else
                            RemoveItem = RemoveItem(value, lst)
                        End If

                    Catch ex As Exception
                        RemoveItem = False
                        Control.ReportSubError(listName, "RemoveItem", ex)
                    End Try
                End Sub)
        End Function

        Friend Shared Sub RemoveAllItems(listName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetSelector(listName).Items.Clear()
                    Catch ex As Exception
                        Control.ReportSubError(listName, "RemoveAllItems", ex)
                    End Try
                End Sub)
        End Sub

        Private Shared Function RemoveItem(item As String, lst As Wpf.Primitives.Selector) As Boolean
            Dim i = lst.Items.IndexOf(item)
            If i = -1 Then
                Return False
            Else
                lst.Items.RemoveAt(i)
                Return True
            End If
        End Function

        Friend Shared Function RemoveItemAt(
                         listName As Primitive,
                         index As Primitive
                   ) As Boolean

            App.Invoke(
                Sub()
                    Dim list = GetSelector(listName)

                    If index.IsArray Then
                        RemoveItemAt = False
                        Dim map = index._arrayMap
                        If map Is Nothing OrElse map.Count = 0 Then Return

                        For Each id In map.Values
                            If DoRemove(list, id) Then
                                RemoveItemAt = True
                            End If
                        Next

                    Else
                        RemoveItemAt = DoRemove(list, index)
                    End If
                End Sub)
        End Function

        Private Shared Function DoRemove(
                           list As Wpf.Primitives.Selector,
                           index As Primitive
                    ) As Boolean

            Try
                If Not index.IsNumber Then Return False

                Dim i As Integer = index
                If i > 0 AndAlso i <= list.Items.Count Then
                    list.Items.RemoveAt(i - 1)
                    Return True
                Else
                    Return False
                End If

            Catch ex As Exception
                Control.ReportSubError(list.Name, "RenoveItemAt", ex)
            End Try

            Return False
        End Function

        Friend Shared Function ContainsItem(listName As Primitive, value As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        ContainsItem = GetSelector(listName).Items.Contains(CStr(value))
                    Catch ex As Exception
                        Control.ReportSubError(listName, "ContainsItem", ex)
                    End Try
                End Sub)
        End Function

        Friend Shared Function FindItem(listName As Primitive, value As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        FindItem = 1 + GetSelector(listName).Items.IndexOf(CStr(value))
                    Catch ex As Exception
                        Control.ReportSubError(listName, "FindItem", ex)
                    End Try
                End Sub)
        End Function

        Friend Shared Function FindItemAt(listName As Primitive, value As Primitive, startIndex As Primitive, endIndex As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetSelector(listName)
                        FindItemAt = 0

                        If Not startIndex.IsNumber OrElse Not endIndex.IsNumber Then
                            Return
                        End If

                        Dim items = lst.Items
                        Dim st As Integer = System.Math.Max(0, startIndex - 1)
                        Dim en As Integer = System.Math.Min(endIndex - 1, items.Count - 1)
                        Dim str = CStr(value)

                        For i = st To en Step If(st > en, -1, 1)
                            If CStr(items(i)) = str Then
                                FindItemAt = i + 1
                                Return
                            End If
                        Next

                    Catch ex As Exception
                        Control.ReportSubError(listName, "FindItemAt", ex)
                    End Try
                End Sub)
        End Function

    End Class
End Namespace
