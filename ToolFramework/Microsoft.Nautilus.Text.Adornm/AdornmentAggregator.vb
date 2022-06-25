Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Core
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.AdornmentSystem
    Friend Class AdornmentAggregator
        Implements IAdornmentProvider

        Friend _adornmentProviders As List(Of IAdornmentProvider)
        Private _adornmentProviderFactories As IEnumerable(Of ImportInfo(Of Func(Of ITextView, IAdornmentProvider), IContentTypeMetadata))
        Private _textView As ITextView

        Public Event AdornmentsChanged As EventHandler(Of AdornmentsChangedEventArgs) Implements IAdornmentProvider.AdornmentsChanged

        Friend Sub New(textView As ITextView, providers As IEnumerable(Of ImportInfo(Of Func(Of ITextView, IAdornmentProvider), IContentTypeMetadata)))
            _textView = textView
            _adornmentProviderFactories = providers
            _adornmentProviders = New List(Of IAdornmentProvider)
            BuildAdornmentProviderList()
            AddHandler textView.TextBuffer.ContentTypeChanged,
                                 Sub() BuildAdornmentProviderList()
        End Sub

        Public Function GetAdornments(span As SnapshotSpan) As IList(Of IAdornment) Implements IAdornmentProvider.GetAdornments
            Dim list1 As New List(Of IAdornment)
            For Each adornmentProvider As IAdornmentProvider In _adornmentProviders
                list1.AddRange(adornmentProvider.GetAdornments(span))
            Next
            Return list1
        End Function

        Private Sub OnAdornmentsChanged(sender As Object, e As AdornmentsChangedEventArgs)
            RaiseEvent AdornmentsChanged(sender, e)
        End Sub

        Private Sub BuildAdornmentProviderList()
            For Each adornmentProvider2 As IAdornmentProvider In _adornmentProviders
                RemoveHandler adornmentProvider2.AdornmentsChanged, AddressOf OnAdornmentsChanged
            Next

            _adornmentProviders.Clear()
            For Each adornmentProviderFactory As ImportInfo(Of Func(Of ITextView, IAdornmentProvider), IContentTypeMetadata) In _adornmentProviderFactories
                For Each contentType As String In adornmentProviderFactory.Metadata.ContentTypes
                    If ContentTypeHelper.IsOfType(_textView.TextBuffer.ContentType, contentType) Then
                        Dim adornmentProvider = adornmentProviderFactory.GetBoundValue()(_textView)
                        If adornmentProvider Is Nothing Then
                            Throw New InvalidOperationException("Adornment Providers cannot be null")
                        End If

                        _adornmentProviders.Add(adornmentProvider)
                        AddHandler adornmentProvider.AdornmentsChanged, AddressOf OnAdornmentsChanged
                        Exit For
                    End If
                Next
            Next
        End Sub
    End Class
End Namespace
