Imports System.Windows.Resources
Imports System.Xml
Imports System.IO
Imports System.Windows.Markup
Imports System.Windows.Threading

Public Class Helper

    Friend Const MmToPx = 96 / 25.4
    Friend Const CmToPx = 96 / 2.54
    Friend Const PxToCm = 2.54 / 96

    Shared Function GetDiagramPanel(element As DependencyObject) As DiagramPanel
        If element Is Nothing Then Return Nothing
        Dim Parent = VisualTreeHelper.GetParent(element)
        Do
            If Parent Is Nothing OrElse TypeOf (Parent) Is DiagramPanel Then Return Parent
            Parent = VisualTreeHelper.GetParent(Parent)
        Loop
    End Function

    Shared Function GetParent(Of ParentType)(child As DependencyObject) As ParentType
        If child Is Nothing Then Return Nothing
        Dim p = child
        Dim t = GetType(ParentType)
        Do
            p = VisualTreeHelper.GetParent(p)
            If p Is Nothing Then Return Nothing
            If p.GetType Is t Then Return CType(CObj(p), ParentType)
        Loop
    End Function

    Shared Function GetListBoxItem(element As DependencyObject) As ListBoxItem
        If element Is Nothing Then Return Nothing
        Dim Parent = VisualTreeHelper.GetParent(element)
        Do
            If Parent Is Nothing OrElse TypeOf (Parent) Is ListBoxItem Then Return Parent
            Parent = VisualTreeHelper.GetParent(Parent)
        Loop
    End Function

    Shared Function GetDesigner(element As DependencyObject) As Designer
        Dim Parent = VisualTreeHelper.GetParent(element)
        Do
            If Parent Is Nothing OrElse TypeOf (Parent) Is Designer Then Return Parent
            Parent = VisualTreeHelper.GetParent(Parent)
        Loop
    End Function

    Shared Function GetCanvas(element As DependencyObject) As Canvas
        Dim Parent = VisualTreeHelper.GetParent(element)
        Do
            If Parent Is Nothing OrElse TypeOf (Parent) Is Canvas Then Return Parent
            Parent = VisualTreeHelper.GetParent(Parent)
        Loop
    End Function

    Shared Function GetScrollViewer(element As DependencyObject) As ScrollViewer
        Dim Parent = VisualTreeHelper.GetParent(element)
        Do
            If Parent Is Nothing OrElse TypeOf (Parent) Is ScrollViewer Then Return Parent
            Parent = VisualTreeHelper.GetParent(Parent)
        Loop
    End Function

    Friend Shared Sub Rotate(obj As FrameworkElement, angle As Double)
        Dim RotateTransform As New RotateTransform
        RotateTransform.Angle = angle
        obj.RenderTransformOrigin = New Point(0.5, 0.5)
        obj.RenderTransform = RotateTransform
        obj.InvalidateMeasure()
    End Sub

    Shared Function GetDiagram(Pnl As DiagramPanel) As FrameworkElement
        Return TryCast(Pnl.Content, ContentPresenter)?.Content
    End Function

    Shared Sub Delay(Milliseconds As Integer)
        Dim T1 = Now
        Do
            If Now.Subtract(T1).TotalMilliseconds >= Milliseconds Then Exit Do
        Loop
    End Sub

    Shared Sub GetTreeElemets(d As DependencyObject, ByRef Item As ListBoxItem, ByRef Pnl As DiagramPanel, ByRef Diagram As UIElement)
        Item = TryCast(d, ListBoxItem)
        If Item IsNot Nothing Then
            Pnl = VisualTreeHelper.GetChild(Item, 0)
            Diagram = Item.Content
        Else
            Pnl = TryCast(d, DiagramPanel)
            If Pnl IsNot Nothing Then
                Item = Pnl.DesignerItem
                Diagram = Helper.GetDiagram(Pnl)
            Else
                Diagram = TryCast(d, UIElement)
                If Diagram IsNot Nothing Then
                    Pnl = Helper.GetListDiagramPanel(Diagram)
                    Item = Pnl.DesignerItem
                End If
            End If
        End If


    End Sub

    Private Shared Function GetListDiagramPanel(Diagram As UIElement) As DiagramPanel
        Dim Parent = VisualTreeHelper.GetParent(Diagram)
        Do
            If Parent Is Nothing OrElse TypeOf (Parent) Is DiagramPanel Then Return Parent
            Parent = VisualTreeHelper.GetParent(Parent)
        Loop
    End Function

    Shared Function FixToMm(Value As Double) As Double
        If Value = 0 OrElse Double.IsNaN(Value) Then Return Value
        Dim mm As Integer = Math.Round(Value * 25.4 / 96)
        Return mm * 96 / 25.4
    End Function

    Public Shared Function LoadXaml(xamlPath As String) As Object
        Dim stream = IO.File.Open(xamlPath, IO.FileMode.Open)
        Dim obj As Object = Nothing
        Try
            obj = XamlReader.Load(stream)
        Catch
        End Try
        Return obj
    End Function

    Shared Function CreateDiagram(FileName As String) As UIElement
        Return CType(LoadXaml(FileName), UIElement)
    End Function

    Shared Function CreateWrapPanel(FileName As String) As WrapPanel
        Dim Rd = CType(LoadXaml(FileName), ResourceDictionary)
        Return Rd("DiagramWrapPanel")
    End Function

    Shared Function Clone(Element As Object) As Object
        Dim xaml = XamlWriter.Save(Element)
        Return XamlReader.Load(XmlReader.Create(New StringReader(xaml)))
    End Function

    Shared Sub UpdateControl(DisObj As DispatcherObject)
        If DisObj Is Nothing Then Return
        Try
            DisObj.Dispatcher.Invoke(DispatcherPriority.Render, Sub() Exit Sub)
        Catch ex As Exception

        End Try
    End Sub

End Class