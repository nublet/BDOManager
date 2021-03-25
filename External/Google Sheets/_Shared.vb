Namespace External.GoogleSheets

    Public Module _Shared

        Private Const _AppName As String = "BDOManager"
        Public Const SpreadSheetID As String = "1H7mPHFHMF4c_ReMimfIE4Q0-NQ0YyWCNqNzVvurXOJE"

        Private _Scopes() As String = {Google.Apis.Sheets.v4.SheetsService.Scope.Spreadsheets}
        Private _SheetsService As Google.Apis.Sheets.v4.SheetsService = Nothing

        Private Function GetSheetsService() As Google.Apis.Sheets.v4.SheetsService
            Try
                If _SheetsService Is Nothing Then
                    Dim UC As Google.Apis.Auth.OAuth2.UserCredential
                    Using stream As New IO.FileStream("client_secret.json", IO.FileMode.Open, IO.FileAccess.Read)
                        Dim FileName As String = "{0}{1}.json".FormatWith(AprBase.GetBaseDirectory(AprBase.BaseDirectory.AppDataCompanyProduct), _AppName)

                        UC = Google.Apis.Auth.OAuth2.GoogleWebAuthorizationBroker.AuthorizeAsync(Google.Apis.Auth.OAuth2.GoogleClientSecrets.Load(stream).Secrets, _Scopes, "user", Threading.CancellationToken.None, New Google.Apis.Util.Store.FileDataStore(FileName, True)).Result
                    End Using
                    _SheetsService = New Google.Apis.Sheets.v4.SheetsService(New Google.Apis.Services.BaseClientService.Initializer() With {.HttpClientInitializer = UC, .ApplicationName = _AppName})
                End If
            Catch ex As Exception
                ex.ToLog(True)
            End Try

            Return _SheetsService
        End Function

        Public Sub UpdateRange(range As String, useRows As Boolean, values As List(Of List(Of Object)))
            UpdateRange(range, SpreadSheetID, useRows, values)
        End Sub

        Public Sub UpdateRange(range As String, spreadSheetId As String, useRows As Boolean, values As List(Of List(Of Object)))
            Try
                Dim ValueRange As New Google.Apis.Sheets.v4.Data.ValueRange()

                If useRows Then
                    ValueRange.MajorDimension = "ROWS"
                Else
                    ValueRange.MajorDimension = "COLUMNS"
                End If

                ValueRange.Values = New List(Of IList(Of Object))() From {}
                For Each Current As List(Of Object) In values
                    ValueRange.Values.Add(Current)
                Next

                Dim Update As Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.UpdateRequest = GetSheetsService().Spreadsheets.Values.Update(ValueRange, spreadSheetId, range)
                Update.ValueInputOption = Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED
                Update.Execute()

                System.Threading.Thread.Sleep(100)
            Catch ex As Exception
                ex.ToLog(True)
            End Try
        End Sub

    End Module

End Namespace