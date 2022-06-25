Imports System.ComponentModel.Composition

Namespace Microsoft.Nautilus.Text.AdornmentSystem

    <Export(GetType(IAdornmentSurfaceManagerFactory))>
    Public NotInheritable Class AdornmentSurfaceManagerService
        Implements IAdornmentSurfaceManagerFactory

        <Import>
        Public Property SurfaceSelector As IAdornmentSurfaceSelector

        Public Function CreateAdornmentSurfaceManager(surfaceHost As IAdornmentSurfaceHost) As IAdornmentSurfaceManager Implements IAdornmentSurfaceManagerFactory.CreateAdornmentSurfaceManager
            If surfaceHost Is Nothing Then
                Throw New ArgumentNullException("surfaceHost")
            End If

            If surfaceHost.TextView Is Nothing Then
                Throw New ArgumentException("The specified surfaceHost isn't valid.")
            End If

            Dim typeFromHandle As Type = GetType(AdornmentSurfaceManagerService)
            Dim [property] As AdornmentSurfaceManager = Nothing
            If Not surfaceHost.TextView.Properties.TryGetProperty(typeFromHandle, [property]) Then
                [property] = New AdornmentSurfaceManager(surfaceHost, _surfaceSelector)
                surfaceHost.TextView.Properties.AddProperty(typeFromHandle, [property])
            End If
            Return [property]
        End Function
    End Class
End Namespace
