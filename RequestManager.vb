Imports System.IO
Imports Newtonsoft.Json

Public Module RequestManager

    Private _JSONErrors As New Dictionary(Of String, List(Of String))
    Private _LastURL As String = ""

    Public Property CacheTime As Integer = 1

    Public ReadOnly Property JSONErrors As Dictionary(Of String, List(Of String))
        Get
            Return _JSONErrors
        End Get
    End Property

    Private Function GetJSONString(cacheFor As Integer, fileName As String, url As String) As String
        url.DownloadToFile(cacheFor, fileName, 5)

        If Not fileName.FileExists() Then
            Return ""
        End If

        Using Reader As New IO.StreamReader(fileName)
            Return Reader.ReadToEnd()
        End Using
    End Function

    Public Function GetEntity(Of T)(url As String) As T
        Return GetEntity(Of T)(CacheTime, url)
    End Function

    Public Function GetEntity(Of T)(cacheFor As Integer, url As String) As T
        Dim FileName As String = url

        If FileName.StartsWith("https://raw.githubusercontent.com/Flockenberger/LifeBDO/master/LifeBDO/") Then
            FileName = "{0}LifeBDO\{1}".FormatWith(JSONCacheLocation, FileName.Substring(FileName.LastIndexOf("/") + 1))
        End If
        If FileName.StartsWith("https://www.bdodae.com/") Then
            FileName = FileName.Substring(23)
            FileName = FileName.Replace("/index.php?cat=", "\").Replace("/", "\")

            FileName = "{0}BDODAE\{1}.json".FormatWith(JSONCacheLocation, FileName)
        End If
        If FileName.StartsWith("https://omegapepega.com/") Then
            FileName = FileName.Substring(27)
            FileName = FileName.Substring(0, FileName.IndexOf("/"))

            Dim ItemID As Integer = AprBase.Type.ToIntegerDB(FileName)

            FileName = "{0}OmegaPepega\{1}000\{2}.json".FormatWith(JSONCacheLocation, (ItemID \ 1000), ItemID)
        End If
        If FileName.StartsWith("https://raw.githubusercontent.com/kookehs/bdo-marketplace/") Then
            FileName = "{0}BDOMarketplace\{1}".FormatWith(JSONCacheLocation, FileName.Substring(FileName.LastIndexOf("/") + 1))
        End If

        If FileName.ToLower().StartsWith("http") Then
            AprBase.Extensions.System_String.ToLog("RequestManager -> GetEntity. Base URL: {0}.".FormatWith(url), True)
            Return Nothing
        End If

        CheckDirectoryExists(IO.Path.GetDirectoryName(FileName))

        Return GetEntity(Of T)(cacheFor, FileName, url)
    End Function

    Public Function GetEntity(Of T)(cacheFor As Integer, fileName As String, url As String) As T
        _LastURL = url

        Return JsonConvert.DeserializeObject(Of T)(GetJSONString(cacheFor, fileName, url), New JsonSerializerSettings With {.MissingMemberHandling = MissingMemberHandling.Error, .Error = AddressOf JSON_Error})
    End Function

    Private Sub CheckDirectoryExists(folderName As String)
        If folderName.StartsWith("\\") Then
            Return
        End If

        Dim CurrentFolder As String = ""

        For Each partName As String In folderName.Split("\"c)
            CurrentFolder.Append("\", partName)

            If CurrentFolder.EndsWith(":") Then
                Continue For
            End If

            If IO.Directory.Exists(CurrentFolder) Then
                Continue For
            End If

            IO.Directory.CreateDirectory(CurrentFolder)
        Next
    End Sub

    Private Sub JSON_Error(sender As Object, e As Serialization.ErrorEventArgs)
        Dim Key As String = e.ErrorContext.Error.Message
        Dim JSONPath As String = ""

        If Key.Contains(". Path ") Then
            JSONPath = Key.Substring(Key.IndexOf(". Path ") + 2).Trim()
            Key = Key.Substring(0, Key.IndexOf(". Path ") + 1)
        End If

        If Not JSONErrors.ContainsKey(Key) Then
            JSONErrors.Add(Key, New List(Of String))
        End If

        If JSONPath.IsSet() Then
            JSONErrors(Key).Add("{0} - {1}".FormatWith(JSONPath, _LastURL))
        Else
            JSONErrors(Key).Add(_LastURL)
        End If

        e.ErrorContext.Handled = True
    End Sub

End Module