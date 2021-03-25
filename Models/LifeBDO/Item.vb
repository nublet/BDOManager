Imports Newtonsoft.Json

Namespace Models.LifeBDO

    Public Class Item

        <JsonProperty("id")> Public Property id As Integer
        <JsonProperty("name")> Public Property name As String
        <JsonProperty("grade")> Public Property grade As Integer
        <JsonProperty("icon")> Public Property icon As String

    End Class

End Namespace