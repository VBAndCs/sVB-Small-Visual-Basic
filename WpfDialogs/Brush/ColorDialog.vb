Public Class ColorDialog

    Private Shared _Brush As Brush

    Public Shared ReadOnly Property Brush() As Brush
        Get
            Return _Brush
        End Get
    End Property

    Public Shared Function Show(SelectedBrush As Brush) As Boolean
        Dim Cd As New WndColor
        Cd.Brush = HatchBrushes.Clone(SelectedBrush)
        If Cd.ShowDialog() Then
            If Cd.Brush IsNot Nothing Then
                _Brush = Cd.Brush
            Else
                _Brush = Nothing
            End If

            Return True
        Else
            Return False
        End If
    End Function

End Class
