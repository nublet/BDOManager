Imports Newtonsoft.Json

Namespace Models.LifeBDO

    Public Class ItemList

        <JsonProperty("items")> Public Property items As List(Of Item)

    End Class

End Namespace