﻿Imports System.Collections.ObjectModel
Imports System.Windows.Input

Namespace Microsoft.SmallVisualBasic.Shell
    Public Class CommandRegistry
        Inherits ObservableCollection(Of ICommand)
    End Class
End Namespace
