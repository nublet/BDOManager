Imports Newtonsoft.Json

Namespace Models.LifeBDO

    Public Class RecipeList

        <JsonProperty("recipes")> Public Property recipes As List(Of Recipe)

    End Class

End Namespace