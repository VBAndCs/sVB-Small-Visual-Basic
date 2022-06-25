Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.AdornmentSystem
    <Export(GetType(IAdornmentSurfaceSelector))>
    Public NotInheritable Class AdornmentSurfaceService
        Implements IAdornmentSurfaceSelector

        <Import(GetType(AdornmentSurfaceFactory))>
        Public Property AdornmentSurfaceFactories As IEnumerable(Of ImportInfo(Of Func(Of IAvalonTextView, IAdornmentSurface), IAdornmentSurfaceFactoryMetadata))


        Public Function CreateAdornmentSurface(textView As IAvalonTextView, adornmentType As Type) As IAdornmentSurface Implements IAdornmentSurfaceSelector.CreateAdornmentSurface
            Return StaticCreateAdornmentSurface(textView, adornmentType, _adornmentSurfaceFactories)
        End Function

        Friend Shared Function StaticCreateAdornmentSurface(textView As IAvalonTextView, adornmentType As Type, adornmentSurfaceFactories1 As IEnumerable(Of ImportInfo(Of Func(Of IAvalonTextView, IAdornmentSurface), IAdornmentSurfaceFactoryMetadata))) As IAdornmentSurface
            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            If adornmentType Is Nothing Then
                Throw New ArgumentNullException("adornmentType")
            End If

            If adornmentSurfaceFactories1 Is Nothing Then
                Throw New ArgumentNullException("adornmentSurfaceFactories")
            End If

            For Each adornmentSurfaceFactory1 As ImportInfo(Of Func(Of IAvalonTextView, IAdornmentSurface), IAdornmentSurfaceFactoryMetadata) In adornmentSurfaceFactories1
                If adornmentSurfaceFactory1.Metadata.SupportedAdornmentType = adornmentType.AssemblyQualifiedName Then
                    Return adornmentSurfaceFactory1.GetBoundValue()(textView)
                End If
            Next
            Return Nothing
        End Function
    End Class
End Namespace
