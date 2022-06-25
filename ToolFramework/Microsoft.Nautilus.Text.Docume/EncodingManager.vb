Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports System.IO
Imports System.Text
Imports Microsoft.Contracts
Imports Microsoft.Nautilus.Core

Namespace Microsoft.Nautilus.Text.Document

    <Export(GetType(IEncodingManager))>
    Public Class EncodingManager
        Implements IEncodingManager

        Private Class NonStreamClosingStreamReader
            Inherits StreamReader

            Friend Sub New(stream As Stream, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean)
                MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks)
            End Sub

            Protected Overrides Sub Dispose(disposing As Boolean)
                MyBase.Dispose(disposing:=False)
            End Sub
        End Class

        Private _orderedSniffers As IList(Of ImportInfo(Of IEncodingSniffer, IOrderable))

        <Import>
        Public Property Sniffers As IEnumerable(Of ImportInfo(Of IEncodingSniffer, IOrderable))

        <Import>
        Public Property Orderer As IOrderer

        Private ReadOnly Property OrderedSniffers As IList(Of ImportInfo(Of IEncodingSniffer, IOrderable))
            Get
                If _orderedSniffers Is Nothing Then
                    _orderedSniffers = New List(Of ImportInfo(Of IEncodingSniffer, IOrderable))(Orderer.Order(Sniffers))
                End If

                Return Me._orderedSniffers
            End Get
        End Property

        Public Function GetSniffedEncoding(stream As Stream) As Encoding Implements IEncodingManager.GetSniffedEncoding
            Contract.RequiresNotNull(stream, "stream")
            Dim pos = stream.Position
            For Each orderedSniffer In OrderedSniffers
                Dim streamEncoding As Encoding = orderedSniffer.GetBoundValue().GetStreamEncoding(stream)
                stream.Position = pos
                If streamEncoding IsNot Nothing Then
                    Return streamEncoding
                End If
            Next
            Return Nothing
        End Function
    End Class
End Namespace
