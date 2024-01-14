Imports WpfDialogs
Imports System.IO
Imports System.Windows.Markup
Imports System.Xml

Class ValuePair(Of T)
    Public OldValue As T
    Public NewValue As T

    Sub New(OldValue As T)
        Me.OldValue = OldValue
    End Sub

    Sub New(OldValue As T, NewValue As T)
        Me.OldValue = OldValue
        Me.NewValue = NewValue
    End Sub
End Class

Class PropertyState
    Inherits Dictionary(Of DependencyProperty, ValuePair(Of Object))
    Implements IRestore

    Public AfterRestoreAction As Action
    Public Owner As DependencyObject

    Sub New(Owner As DependencyObject)
        Me.Owner = Owner
    End Sub

    Sub New(ParamArray Prop() As DependencyProperty)
        Me.Add(Prop)
    End Sub

    Sub New(Owner As DependencyObject, ParamArray Prop() As DependencyProperty)
        Me.Owner = Owner
        Me.Add(Prop)
    End Sub

    Sub New(AfterRestoreAction As Action, Owner As DependencyObject, ParamArray Prop() As DependencyProperty)
        Me.AfterRestoreAction = AfterRestoreAction
        Me.Owner = Owner
        Me.Add(Prop)
    End Sub

    Private Function Clone(Element As Object) As Object
        If Element Is Nothing Then Return Nothing
        If TypeOf Element Is ValueType OrElse TypeOf Element Is String Then Return Element

        Dim xaml = XamlWriter.Save(Element)
        Return XamlReader.Load(XmlReader.Create(New StringReader(xaml)))
    End Function

    Private Function AreEqual(Element1 As Object, Element2 As Object) As Boolean
        Dim xaml1 = If(Element1 Is Nothing, "", XamlWriter.Save(Element1))
        Dim xaml2 = If(Element2 Is Nothing, "", XamlWriter.Save(Element2))
        Return xaml1 = xaml2
    End Function

    Overloads Sub Add(ParamArray Prop() As DependencyProperty)
        For Each P In Prop
            Me.Add(P, New ValuePair(Of Object)(Me.Clone(Owner.GetValue(P))))
        Next
    End Sub

    Function HasChanges() As Boolean
        For Each Pair In Me
            If Not Me.AreEqual(Pair.Value.OldValue, Owner.GetValue(Pair.Key)) Then Return True
        Next
        Return False
    End Function

    Sub RestoreOldValues() Implements IRestore.RestoreOldValues
        For Each Pair In Me
            Owner.SetValue(Pair.Key, Me.Clone(Pair.Value.OldValue))
        Next
        If AfterRestoreAction IsNot Nothing Then AfterRestoreAction()
    End Sub

    Sub RestoreNewValues() Implements IRestore.RestoreNewValues
        For Each Pair In Me
            Owner.SetValue(Pair.Key, Me.Clone(Pair.Value.NewValue))
        Next
        If AfterRestoreAction IsNot Nothing Then AfterRestoreAction()
    End Sub

    Function SetNewValues() As PropertyState
        For Each Pair In Me
            Pair.Value.NewValue = Me.Clone(Owner.GetValue(Pair.Key))
        Next
        Return Me
    End Function

End Class
