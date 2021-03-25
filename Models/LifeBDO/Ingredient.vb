Imports Newtonsoft.Json

Namespace Models.LifeBDO

    Public Class Ingredient

        <JsonProperty("item")> Public Property item As List(Of Item)
        <JsonProperty("amount")> Public Property amount As Integer

    End Class

End Namespace