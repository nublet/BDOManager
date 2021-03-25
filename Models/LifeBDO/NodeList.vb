Imports Newtonsoft.Json

Namespace Models.LifeBDO

    Public Class NodeList

        <JsonProperty("nodes")> Public Property nodes As List(Of Node)

    End Class

End Namespace