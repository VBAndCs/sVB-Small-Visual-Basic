Imports System
Imports System.Windows
Imports System.Windows.Input

Namespace Microsoft.SmallVisualBasic.Shell
    Public Class RoutedRibbonCommand
        Inherits DependencyObject
        Implements ICommand

        Private nameField As String
        Public Shared ReadOnly DescriptionProperty As DependencyProperty = DependencyProperty.Register("Description", GetType(String), GetType(RoutedRibbonCommand))
        Public Property Command As RoutedCommand

        Public Property Name As String
            Get

                If ShowCommandText Then
                    If Equals(nameField, Nothing) Then
                        Dim routedUICommand As RoutedUICommand = TryCast(Command, RoutedUICommand)

                        If routedUICommand IsNot Nothing Then
                            Return routedUICommand.Text
                        End If

                        Return Command.Name
                    End If

                    Return nameField
                End If

                Return Nothing
            End Get
            Set(value As String)
                nameField = value
            End Set
        End Property

        Public Property Description As String
            Get
                Return CStr(GetValue(DescriptionProperty))
            End Get
            Set(value As String)
                SetValue(DescriptionProperty, value)
            End Set
        End Property

        Public Property ShowCommandText As Boolean

        Public Custom Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged
            AddHandler(handler As EventHandler)
                AddHandler Command.CanExecuteChanged, handler
            End AddHandler

            RemoveHandler(handler As EventHandler)
                RemoveHandler Command.CanExecuteChanged, handler
            End RemoveHandler

            RaiseEvent(sender As Object, e As EventArgs)
            End RaiseEvent
        End Event

        Public Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
            Return Command.CanExecute(parameter, Keyboard.FocusedElement)
        End Function

        Public Sub Execute(parameter As Object) Implements ICommand.Execute
            Command.Execute(parameter, Keyboard.FocusedElement)
        End Sub
    End Class
End Namespace
