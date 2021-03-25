Imports Newtonsoft.Json

Namespace Models.OmegaPepega

    Public Class Item

        <JsonProperty("chooseKey")> Public Property chooseKey As Integer
        <JsonProperty("count")> Public Property count As Integer
        <JsonProperty("grade")> Public Property grade As Integer
        <JsonProperty("keyType")> Public Property keyType As Integer
        <JsonProperty("mainCategory")> Public Property mainCategory As Integer
        <JsonProperty("mainKey")> Public Property mainKey As Integer
        <JsonProperty("name")> Public Property name As String
        <JsonProperty("pricePerOne")> Public Property pricePerOne As Long
        <JsonProperty("subCategory")> Public Property subCategory As Integer
        <JsonProperty("subKey")> Public Property subKey As Integer
        <JsonProperty("totalTradeCount")> Public Property totalTradeCount As Long

    End Class

End Namespace