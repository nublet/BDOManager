Imports Newtonsoft.Json

Namespace Models.LifeBDO

    Public Class Recipe

        <JsonProperty("product")> Public Property product As Item
        <JsonProperty("lifeSkill")> Public Property lifeSkill As String
        <JsonProperty("level")> Public Property level As String
        <JsonProperty("xp")> Public Property xp As Integer
        <JsonProperty("effects")> Public Property effects As List(Of Effect)
        <JsonProperty("ingredients")> Public Property ingredients As List(Of Ingredient)

    End Class

End Namespace