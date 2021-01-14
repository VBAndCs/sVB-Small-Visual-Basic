Imports System.ComponentModel

Public Class PropertyData
    Public Property OwnerType As Object

    Dim _Value As Object

    Public Property Value As Object
        Get
            If _Value.ToString = "__NullValue__" Then Return Nothing
            Return _Value
        End Get

        Set(NewValue As Object)
            Dim SafeValue
            If NewValue Is Nothing Then
                SafeValue = "__NullValue__"
            Else
                SafeValue = NewValue
            End If
            _Value = SafeValue
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(OwnerType As Object, Value As Object)
        Me.OwnerType = OwnerType
        Me.Value = Value
    End Sub

    Public Sub New(Prop As DependencyProperty, Value As Object)
        Me.OwnerType = Prop.OwnerType
        Me.Value = Value
    End Sub
End Class

Public Class PropertyDictionary
    Inherits Dictionary(Of String, PropertyData)

    Dim DependencyObject As DependencyObject

    Public Sub New()

    End Sub

    Public Sub New(Obj As DependencyObject, ParamArray Props() As DependencyProperty)
        Me.DependencyObject = Obj
        For Each Prop In Props
            MyBase.Add(Prop.Name, New PropertyData(Prop, Obj.GetValue(Prop)))
        Next
    End Sub

    Public Sub SetDependencyObject(Obj As DependencyObject)
        Me.DependencyObject = Obj
    End Sub

    Public Shadows Sub Add([Property] As DependencyProperty, Value As Object)
        Me([Property]) = Value
    End Sub

    Public Shadows Sub Add(ParamArray Props() As DependencyProperty)
        For Each Prop In Props
            If Not MyBase.ContainsKey(Prop.Name) Then
                MyBase.Add(Prop.Name, New PropertyData(Prop, DependencyObject.GetValue(Prop)))
            End If
        Next
    End Sub

    Default Public Shadows Property Item([Property] As DependencyProperty) As Object
        Get
            Dim PropName = [Property].Name
            If MyBase.ContainsKey(PropName) Then
                Dim PropData As PropertyData = MyBase.Item(PropName)
                Return PropData.Value
            Else
                Return Nothing
            End If
        End Get

        Set(value As Object)
            Dim PropName = [Property].Name
            If Me.ContainsKey(PropName) Then
                MyBase.Item(PropName) = New PropertyData([Property], value)
                DependencyObject.SetValue([Property], value)
            Else
                MyBase.Add(PropName, New PropertyData([Property], value))
                DependencyObject.SetValue([Property], value)
            End If
        End Set
    End Property

    Public Sub SetValuesToObj()
        If Me.DependencyObject Is Nothing Then Return
        For Each Entry In Me
            Dim OwnerType = Entry.Value.OwnerType
            Dim descriptor = DependencyPropertyDescriptor.FromName(Entry.Key, OwnerType, DependencyObject.GetType())
            Dim [Property] = descriptor.DependencyProperty
            DependencyObject.SetValue([Property], Me([Property]))
        Next
    End Sub

    Public Sub UpdateValuesFromObj()
        If Me.DependencyObject Is Nothing Then Return
        For i = 0 To Me.Count - 1
            Dim OwnerType = Me.Values(i).OwnerType
            Dim descriptor = DependencyPropertyDescriptor.FromName(Me.Keys(i), OwnerType, DependencyObject.GetType())
            Dim [Property] = descriptor.DependencyProperty
            Me([Property]) = DependencyObject.GetValue([Property])
        Next
    End Sub
End Class
