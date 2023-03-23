
Imports System.Text
Imports System.Threading
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' This class provides access to Flickr photo services.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class Flickr
        Private Const _apiKey As String = "MWY5ZmI5ODE4Mjk2NzAwNTgwYmYzMzQwMjc5MzQ2YjY="

        ''' <summary>
        ''' Flickr URL Template (with https)
        ''' </summary>
        Private Shared _urlTemplate As String
        Private Shared _picUrlTemplate As String
        Private Shared _cachedTag As String
        Private Shared _cachedTagDocument As XDocument
        Private Shared _cachedTagDocumentHitCount As Integer

        Shared Sub New()
            _urlTemplate = "https://api.flickr.com/services/rest/?sort=interestingness-desc&safe_search=1&license=1,2,3,4,5,6,7&api_key="
            _picUrlTemplate = "http://farm{0}.static.flickr.com/{1}/{2}_{3}.jpg"
            Dim array As Byte() = Convert.FromBase64String("MWY5ZmI5ODE4Mjk2NzAwNTgwYmYzMzQwMjc5MzQ2YjY=")
            _urlTemplate &= Encoding.UTF8.GetString(array, 0, array.Length)
        End Sub

        ''' <summary>
        ''' Gets the URL for the picture of the moment.
        ''' </summary>
        ''' <returns>
        ''' A file URL for Flickr's picture of the moment
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetPictureOfMoment() As Primitive
            Dim random1 As New Random(Now.Ticks Mod Integer.MaxValue)
            Dim url As String = $"{_urlTemplate}&method=flickr.photos.getRecent&per_page=50&page={random1.Next(10) + 1}"
            Thread.Sleep(1000)
            Return GetRandomPhotoNodeUrl(RestHelper.GetContents(url))
        End Function

        ''' <summary>
        ''' Gets the URL for a random picture tagged with the specified tag.
        ''' </summary>
        ''' <param name="tag">
        ''' The tag for the requested picture.
        ''' </param>
        ''' <returns>
        ''' A file URL for Flickr's random picture
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetRandomPicture(tag As Primitive) As Primitive
            Dim random1 As New Random(Now.Ticks Mod Integer.MaxValue)
            If CStr(tag) = _cachedTag AndAlso _cachedTagDocument IsNot Nothing AndAlso _cachedTagDocumentHitCount < 10 Then
                _cachedTagDocumentHitCount += 1
                Return GetRandomPhotoNodeUrl(_cachedTagDocument)
            End If

            Dim url As String = $"{_urlTemplate}&method=flickr.photos.search&per_page=50&tags={tag}&page={random1.Next(10) + 1}"
            Thread.Sleep(1000)
            _cachedTag = tag
            _cachedTagDocument = RestHelper.GetContents(url)
            _cachedTagDocumentHitCount = 0
            Return GetRandomPhotoNodeUrl(_cachedTagDocument)
        End Function

        Private Shared Function GetRandomPhotoNodeUrl(document As XDocument) As String
            If document Is Nothing Then
                Return ""
            End If
            Try
                Dim source As IEnumerable(Of XElement) = document.Descendants("photo")
                Dim random1 As New Random(Now.Ticks Mod Integer.MaxValue)
                Dim xElement1 As XElement = source.ElementAt(random1.Next(source.Count()))
                Dim id = xElement1.Attribute("id").Value
                Dim secret = xElement1.Attribute("secret").Value
                Dim server = xElement1.Attribute("server").Value
                Dim farm = xElement1.Attribute("farm").Value
                Return String.Format(_picUrlTemplate, farm, server, id, secret)
            Catch

            End Try

            Return ""
        End Function
    End Class
End Namespace
