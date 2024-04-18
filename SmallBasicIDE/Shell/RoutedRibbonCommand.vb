Imports System
Imports System.Windows
Imports System.Windows.Input

Namespace Microsoft.SmallVisualBasic.Shell
    Public Class RoutedRibbonCommand
        Inherits DependencyObject
        Implements ICommand

        Private nameField As String

        Public Shared ReadOnly DescriptionProperty As DependencyProperty = DependencyProperty.Register("Description", GetType(String), GetType(RoutedRibbonCommand))

        Public Shared Commands As New System.Collections.Generic.Dictionary(Of RoutedCommand, RoutedRibbonCommand)

        Dim _command As RoutedCommand
        Public Property Command As RoutedCommand
            Get
                Return _command
            End Get

            Set(value As RoutedCommand)
                _command = value
                Commands(_command) = Me
            End Set
        End Property

        Public Property Name As String
            Get
                If Not ShowCommandText Then Return Nothing

                If Equals(nameField, Nothing) Then
                    Dim routedUICommand As RoutedUICommand = TryCast(Command, RoutedUICommand)
                    If routedUICommand IsNot Nothing Then
                        Return routedUICommand.Text
                    Else
                        Return Command.Name
                    End If
                Else
                    Return nameField
                End If
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
