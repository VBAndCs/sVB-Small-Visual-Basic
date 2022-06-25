Imports System.Collections
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

Namespace Microsoft.Nautilus.Text.Editor
    Public Interface IFormattedTextLineCollection
        Inherits IList(Of ITextLine), ICollection(Of ITextLine), IEnumerable(Of ITextLine), IEnumerable
        ReadOnly Property FirstFullyVisibleLine As ITextLine

        ReadOnly Property LastFullyVisibleLine As ITextLine

        ReadOnly Property FormattedContentWidth As Double

        ReadOnly Property FormattedContentHeight As Double

        Function FindTextLineIndexContainingPosition(position As Integer, <Out> ByRef index As Integer) As Boolean

        Function GetTextLineContainingPosition(position As Integer) As ITextLine

        Function FindTextLineIndexContainingYCoordinate(y As Double, <Out> ByRef index As Integer) As Boolean

        Function GetTextLineContainingYCoordinate(y As Double) As ITextLine

        Function GetIndexOfTextLine(textLine As ITextLine) As Integer

        Function GetVisibilityStateOfTextLine(textLine As ITextLine) As VisibilityState
    End Interface
End Namespace
