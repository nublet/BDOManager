Imports System.Runtime.InteropServices

Namespace Models

    Public Class Recipe

        Public Property EXP As Integer = 0
        Public Property Ingredients As New List(Of Ingredient)
        Public Property LifeSkill As String = ""
        Public Property LifeSkillLevel As String = ""
        Public Property Product As String = ""

        Public Sub New(ByRef template As LifeBDO.Recipe, ByRef validIds As List(Of Integer))
            Me.EXP = template.xp
            Me.LifeSkill = GetName(template.lifeSkill)
            Me.LifeSkillLevel = GetName(template.level)
            Me.Product = GetActualResources(template.product.name).Trim()

            Ingredients.Clear()
            For Each Ingredient As LifeBDO.Ingredient In template.ingredients
                Me.Ingredients.Add(New Models.Ingredient(Ingredient, validIds))
            Next

            validIds.Add(template.product.id)
        End Sub

    End Class

End Namespace