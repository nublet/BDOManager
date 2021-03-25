Imports Newtonsoft.Json

Namespace Models.BDOMarketplace

    Public Class Item

        <JsonProperty("id")> Public Property id As Integer
        <JsonProperty("name")> Public Property name As String
        <JsonProperty("grade")> Public Property grade As Integer

    End Class

End Namespace