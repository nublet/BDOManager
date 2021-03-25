Imports Newtonsoft.Json

Namespace Models.LifeBDO

    Public Class Effect

        <JsonProperty("effect")> Public Property effect As String
        <JsonProperty("amplifier")> Public Property amplifier As Integer
        <JsonProperty("percent")> Public Property percent As Boolean

    End Class

End Namespace