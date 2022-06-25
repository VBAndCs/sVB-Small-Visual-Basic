Imports System.Collections.Generic
Imports System.Windows.Media
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.AdornmentSystem
    Friend Class AdornmentSurfaceManager
        Implements IAdornmentSurfaceManager
        Implements IAdornmentSurfaceSpaceManager

        Private _surfaceHost As IAdornmentSurfaceHost
        Friend _adornmentSurfaces As Dictionary(Of Type, IAdornmentSurface)
        Private _surfaceSelector As IAdornmentSurfaceSelector

        Friend Sub New(
                      surfaceHost As IAdornmentSurfaceHost,
                      surfaceSelector As IAdornmentSurfaceSelector)

            _surfaceHost = surfaceHost
            _surfaceSelector = surfaceSelector
            _adornmentSurfaces = New Dictionary(Of Type, IAdornmentSurface)
        End Sub

        Public Sub AddAdornment(adornment As IAdornment) Implements IAdornmentSurfaceManager.AddAdornment
            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If

            GetAdornmentSurface(adornment.GetType())?.AddAdornment(adornment)
        End Sub

        Public Sub RemoveAdornment(adornment As IAdornment) Implements IAdornmentSurfaceManager.RemoveAdornment
            If adornment Is Nothing Then
                Throw New ArgumentNullException("adornment")
            End If

            GetAdornmentSurface(adornment.GetType())?.RemoveAdornment(adornment)
        End Sub

        Public Function GetSpaceNegotiations(textLine As ITextLine) As IList(Of SpaceNegotiation) Implements IAdornmentSurfaceManager.GetSpaceNegotiations
            If textLine Is Nothing Then
                Throw New ArgumentNullException("textLine")
            End If

            Dim list1 As New List(Of SpaceNegotiation)
            For Each _value As IAdornmentSurface In _adornmentSurfaces.Values
                If _value IsNot Nothing Then
                    Dim spaceNegotiations As IList(Of SpaceNegotiation) = _value.GetSpaceNegotiations(textLine)
                    If spaceNegotiations IsNot Nothing Then
                        list1.AddRange(spaceNegotiations)
                    End If
                End If
            Next
            Return list1
        End Function

        Public Function GetVisibleAdornmentsGeometry() As Geometry Implements IAdornmentSurfaceSpaceManager.GetVisibleAdornmentsGeometry
            Dim pathGeometry1 As New PathGeometry
            pathGeometry1.FillRule = FillRule.Nonzero
            For Each _value As IAdornmentSurface In _adornmentSurfaces.Values
                If _value IsNot Nothing AndAlso _value.CanNegotiateSurfaceSpace Then
                    Dim adornmentsGeometry As Geometry = _value.GetAdornmentsGeometry()
                    If adornmentsGeometry IsNot Nothing Then
                        pathGeometry1.AddGeometry(adornmentsGeometry)
                    End If
                End If
            Next
            Return pathGeometry1.GetOutlinedPathGeometry()
        End Function

        Private Function GetAdornmentSurface(adornmentType As Type) As IAdornmentSurface
            Dim _value As IAdornmentSurface = Nothing
            If _adornmentSurfaces.TryGetValue(adornmentType, _value) Then
                Return _value
            End If

            _value = _surfaceSelector.CreateAdornmentSurface(_surfaceHost.TextView, adornmentType)
            _adornmentSurfaces(adornmentType) = _value
            If _value IsNot Nothing Then
                _surfaceHost.AddAdornmentSurface(_value)
            End If
            Return _value
        End Function
    End Class
End Namespace
