Imports Newtonsoft.Json

Namespace Models.LifeBDO

    Public Class Node

        <JsonProperty("node")> Public Property node As String
        <JsonProperty("region")> Public Property region As String
        <JsonProperty("cpCost")> Public Property cpCost As Integer
        <JsonProperty("subNodes")> Public Property subNodes As List(Of NodeResource)

    End Class

End Namespace