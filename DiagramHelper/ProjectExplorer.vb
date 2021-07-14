Imports System.Windows.Controls.Primitives
Imports System.ComponentModel

Public Class ProjectExplorer
    Inherits Control

    Public Sub New()
        Dim resourceLocater As Uri = New Uri("/DiagramHelper;component/Resources/projectexplorerdecorator.xaml", System.UriKind.Relative)
        Dim ResDec As ResourceDictionary = Application.LoadComponent(resourceLocater)
        Me.Resources.MergedDictionaries.Add(ResDec)
        Me.Style = FindResource("ProjectExplorerStyle")
    End Sub

    Public WithEvents FilesList As ListBox

    Dim FirstTime As Boolean = True

    Private Sub ProjectExplorer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If FirstTime Then
            FirstTime = False

            FilesList = TryCast(Template.FindName("PART_ListBox", Me), ListBox)
            FilesList.ItemsSource = Designer.FormNames
            FilesList.SelectedIndex = 0
        End If
    End Sub

    Public FreezListFiles As Boolean

    Sub FilesList_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles FilesList.SelectionChanged
        If FreezListFiles Then Return
        Dim i = FilesList.SelectedIndex
        If i = -1 Then Return

        Designer.SwitchTo(Designer.FormKeys(i))
    End Sub

    Private Sub FilesList_KeyDown(sender As Object, e As KeyEventArgs) Handles FilesList.KeyDown
        If e.Key = Key.Delete Then
            If Designer.IsNew AndAlso Designer.Pages.Count = 1 Then
                Beep()
                Return
            End If

            If FilesList.SelectedIndex = -1 Then Return
            Designer.ClosePage()
        End If
    End Sub

    Public Property Designer As Designer
        Get
            Return GetValue(DesignerProperty)
        End Get

        Set(ByVal value As Designer)
            SetValue(DesignerProperty, value)
        End Set
    End Property

    Public Shared ReadOnly DesignerProperty As DependencyProperty =
                           DependencyProperty.Register("Designer",
                           GetType(Designer), GetType(ProjectExplorer))

End Class
