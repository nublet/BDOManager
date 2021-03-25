Imports Newtonsoft.Json

Namespace Models.LifeBDO

    Public Class NodeResource

        <JsonProperty("cpCost")> Public Property cpCost As Integer
        <JsonProperty("items")> Public Property items As List(Of Item)

    End Class

End Namespace